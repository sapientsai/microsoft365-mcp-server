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

// Bold before italic, so "**x**" isn't read as two italics. Runs on text that is already escaped,
// and the markers (*, _) survive escaping unchanged.
const emphasis = (escaped: string): string =>
  escaped
    .replace(/\*\*(?=\S)(.+?)(?<=\S)\*\*/g, "<b>$1</b>")
    .replace(/(^|[^\w*])\*(?=\S)([^*]+?)(?<=\S)\*(?![\w*])/g, "$1<i>$2</i>")
    .replace(/(^|[^\w])_(?=\S)([^_]+?)(?<=\S)_(?!\w)/g, "$1<i>$2</i>")

// Inline code and links are cut out first, so emphasis never reaches inside them. String.split with
// capture groups gives the plain runs and each token's groups in order: [plain, code, label, href, …].
const INLINE_TOKEN = /`([^`\n]+)`|\[([^\]\n]+)\]\(([^)\s]+)\)/
const INLINE_PARTS = 4

const token = (code: string | undefined, label: string | undefined, href: string | undefined): string => {
  if (code !== undefined) return `<code>${escapeHtml(code)}</code>`
  const url = safeHref(href ?? "")
  return url === undefined
    ? emphasis(escapeHtml(`[${label ?? ""}](${href ?? ""})`))
    : `<a href="${escapeHtml(url)}">${emphasis(escapeHtml(label ?? ""))}</a>`
}

const inline = (text: string): string => {
  const pieces: ReadonlyArray<string | undefined> = text.split(new RegExp(INLINE_TOKEN, "g"))
  return pieces
    .map((piece, k) => {
      const slot = k % INLINE_PARTS
      if (slot === 0) return emphasis(escapeHtml(piece ?? ""))
      return slot === 1 ? token(piece, pieces[k + 1], pieces[k + 2]) : ""
    })
    .join("")
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
const runEnd = (lines: ReadonlyArray<string>, from: number, keep: (line: string) => boolean): number => {
  const stop = lines.findIndex((line, k) => k >= from && !keep(line))
  return stop === -1 ? lines.length : stop
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

// One pass over the lines, without recursion: a long message must not run out of stack.
const blocks = (markdown: string): ReadonlyArray<Block> => {
  const lines = markdown.replace(/\r\n?/g, "\n").split("\n")
  return lines.reduce<{ readonly found: ReadonlyArray<Block>; readonly next: number }>(
    (acc, _line, k) => {
      if (k < acc.next) return acc
      const { block, next } = blockAt(lines, k)
      return { found: block === undefined ? acc.found : [...acc.found, block], next }
    },
    { found: [], next: 0 },
  ).found
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
