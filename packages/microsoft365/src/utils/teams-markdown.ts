// Markdown to the HTML a Teams message body takes. Callers (LLMs) write Markdown already; Teams shows
// a text body as one block, so lists and commands run together. Only a small subset is converted:
// paragraphs and line breaks, **bold**, *italic* / _italic_, `code`, fenced code blocks, "- " and
// "1. " lists, # headings (as a bold line) and [links](https://…).
//
// Every character of the caller's text is escaped, and the only tags in the output are the ones made
// here, so content such as "a < b" or a pasted "<script>" arrives as text, never as markup. A link
// is kept only when it parses as http: or https:, and its href is the parsed url.href.

const escapeHtml = (text: string): string =>
  text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;")

const safeHref = (raw: string): string | undefined => {
  const url = URL.parse(raw)
  return url !== null && (url.protocol === "http:" || url.protocol === "https:") ? url.href : undefined
}

const tagCount = (html: string, tag: string): number => html.split(tag).length - 1

// ***x*** first, then bold, then italic, so "**x**" isn't read as two italics. An italic is kept only
// when the bold tags inside it balance, so "**a *b** c*" can't produce <b>a <i>b</b> c</i>. Runs on
// text that is already escaped; the markers (*, _) survive escaping unchanged.
const emphasis = (escaped: string): string =>
  escaped
    .replace(/\*\*\*(?=\S)(.+?)(?<=\S)\*\*\*/g, "<b><i>$1</i></b>")
    .replace(/\*\*(?=\S)(.+?)(?<=\S)\*\*/g, "<b>$1</b>")
    .replace(/(^|[^\w*])\*(?=\S)([^*]+?)(?<=\S)\*(?![\w*])/g, (whole, lead: string, body: string) =>
      tagCount(body, "<b>") === tagCount(body, "</b>") ? `${lead}<i>${body}</i>` : whole,
    )
    .replace(/(^|[^\w])_(?=\S)([^_]+?)(?<=\S)_(?!\w)/g, (whole, lead: string, body: string) =>
      tagCount(body, "<b>") === tagCount(body, "</b>") ? `${lead}<i>${body}</i>` : whole,
    )

// Inline code and links become placeholders first, so emphasis never reaches inside them but can
// still wrap them ("**run `pnpm i` first**"). A placeholder is U+E000 (private use), its index, U+E000; that character is removed
// from the caller's text beforehand, so no text can forge one.
const INLINE_TOKEN = /`([^`\n]+)`|\[([^\]\n]+)\]\(([^)\s]+)\)/g
const PLACEHOLDER = /\uE000(\d+)\uE000/g

const renderToken = (whole: string, code: string | undefined, label: string, href: string): string => {
  if (code !== undefined) return `<code>${escapeHtml(code)}</code>`
  const url = safeHref(href)
  return url === undefined
    ? emphasis(escapeHtml(whole))
    : `<a href="${escapeHtml(url)}">${emphasis(escapeHtml(label))}</a>`
}

const inline = (raw: string): string => {
  const text = raw.replace(/\uE000/g, "")
  const rendered: string[] = []
  const marked = text.replace(
    INLINE_TOKEN,
    (whole: string, code: string | undefined, label: string | undefined, href: string | undefined, at: number) => {
      // ![alt](src) is an image, which isn't converted: leave all of it as text.
      if (code === undefined && text[at - 1] === "!") return whole
      rendered.push(renderToken(whole, code, label ?? "", href ?? ""))
      return `\uE000${rendered.length - 1}\uE000`
    },
  )
  return emphasis(escapeHtml(marked)).replace(PLACEHOLDER, (_whole, k: string) => rendered[Number(k)] ?? "")
}

type Block =
  | { readonly kind: "code"; readonly lines: ReadonlyArray<string> }
  | { readonly kind: "list"; readonly ordered: boolean; readonly items: ReadonlyArray<string> }
  | { readonly kind: "heading"; readonly text: string }
  | { readonly kind: "paragraph"; readonly lines: ReadonlyArray<string> }

const FENCE = /^\s*(```|~~~)/
const BULLET = /^\s*[-*+]\s+(.*)$/
const NUMBERED = /^\s*\d+[.)]\s+(.*)$/
const HEADING = /^\s{0,3}#{1,6}\s+(.*?)\s*#*\s*$/

const startsBlock = (line: string): boolean =>
  line.trim() === "" || FENCE.test(line) || HEADING.test(line) || BULLET.test(line) || NUMBERED.test(line)

/** The index of the first line at or after `from` that fails `keep`, or the end. */
// A plain loop from `from`: findIndex from 0 on every block made parsing quadratic (30 s for 120,000 lines).
const runEnd = (lines: ReadonlyArray<string>, from: number, keep: (line: string) => boolean): number => {
  let k = from // eslint-disable-line functype/no-let -- linear scan, see above
  // eslint-disable-next-line functype/no-imperative-loops -- linear scan, see above
  while (k < lines.length && keep(lines[k] ?? "")) k++
  return k
}

// One block from `at`, and where the next one starts.
const blockAt = (lines: ReadonlyArray<string>, at: number): { readonly block?: Block; readonly next: number } => {
  const line = lines[at] ?? ""
  const fence = FENCE.exec(line)?.[1]
  if (fence !== undefined) {
    // An unclosed block runs to the end.
    const close = runEnd(lines, at + 1, (l) => !l.trimStart().startsWith(fence))
    return { block: { kind: "code", lines: lines.slice(at + 1, close) }, next: close + 1 }
  }
  if (line.trim() === "") return { next: at + 1 }
  const heading = HEADING.exec(line)?.[1]
  if (heading !== undefined) return { block: { kind: "heading", text: heading }, next: at + 1 }
  const ordered = NUMBERED.test(line)
  if (ordered || BULLET.test(line)) {
    const pattern = ordered ? NUMBERED : BULLET
    const end = runEnd(lines, at, (l) => pattern.test(l))
    const items = lines.slice(at, end).map((l) => pattern.exec(l)?.[1] ?? "")
    return { block: { kind: "list", ordered, items }, next: end }
  }
  const end = runEnd(lines, at + 1, (l) => !startsBlock(l))
  return { block: { kind: "paragraph", lines: lines.slice(at, end) }, next: end }
}

// One linear pass over the lines: no recursion, which could run out of stack, and no copying of the
// blocks found so far, which made it quadratic.
const blocks = (markdown: string): ReadonlyArray<Block> => {
  const lines = markdown.replace(/\r\n?/g, "\n").split("\n")
  const found: Block[] = []
  let at = 0 // eslint-disable-line functype/no-let -- line cursor
  // eslint-disable-next-line functype/no-imperative-loops -- one pass, see above
  while (at < lines.length) {
    const { block, next } = blockAt(lines, at)
    if (block !== undefined) found.push(block)
    at = next
  }
  return found
}

const renderBlock = (block: Block): string => {
  switch (block.kind) {
    case "code":
      return `<pre><code>${escapeHtml(block.lines.join("\n"))}</code></pre>`
    case "heading":
      return `<p><b>${inline(block.text)}</b></p>`
    case "list": {
      const tag = block.ordered ? "ol" : "ul"
      return `<${tag}>${block.items.map((item) => `<li>${inline(item)}</li>`).join("")}</${tag}>`
    }
    case "paragraph":
      return `<p>${block.lines.map(inline).join("<br>")}</p>`
  }
}

/** Markdown to Teams message HTML; see the comment at the top of this file for what is converted. */
export const markdownToTeamsHtml = (markdown: string): string => blocks(markdown).map(renderBlock).join("")

// Teams refuses a message over about 28 KB, so a longer one can't be sent either way, and the inline
// patterns slow down on very long lines. Markdown longer than this is refused before converting.
export const MAX_MARKDOWN_CHARS = 30_000

/** Why content can't be sent in this format, or undefined when it can. */
export const messageFormatProblem = (content: string, format: MessageFormat | undefined): string | undefined =>
  format === "markdown" && content.length > MAX_MARKDOWN_CHARS
    ? `The message is ${content.length.toLocaleString("en-US")} characters, over the ${MAX_MARKDOWN_CHARS.toLocaleString("en-US")} ` +
      "a Markdown message may have. Teams refuses messages over about 28 KB; split it into several messages."
    : undefined

export const MESSAGE_FORMATS = ["text", "markdown"] as const
export type MessageFormat = (typeof MESSAGE_FORMATS)[number]

/** The body Graph takes for a chat or channel message in the given format. */
export const messageBody = (
  content: string,
  format: MessageFormat = "text",
): { readonly contentType: "text" | "html"; readonly content: string } =>
  format === "markdown"
    ? { contentType: "html", content: markdownToTeamsHtml(content) }
    : { contentType: "text", content }
