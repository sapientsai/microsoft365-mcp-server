import { describe, expect, it } from "vitest"

import { MAX_MARKDOWN_CHARS, markdownToTeamsHtml, messageBody, messageFormatProblem } from "../src/utils/teams-markdown"

describe("markdownToTeamsHtml", () => {
  it("turns a fenced code block into pre/code, escaped and unformatted", () => {
    expect(markdownToTeamsHtml("Run:\n\n```bash\nnpm i <pkg> && echo **done**\n```")).toBe(
      "<p>Run:</p><pre><code>npm i &lt;pkg&gt; &amp;&amp; echo **done**</code></pre>",
    )
  })

  it("turns '- ' and '1. ' lines into lists", () => {
    expect(markdownToTeamsHtml("- one\n- **two**")).toBe("<ul><li>one</li><li><b>two</b></li></ul>")
    expect(markdownToTeamsHtml("1. sign in\n2. run `pnpm i`")).toBe(
      "<ol><li>sign in</li><li>run <code>pnpm i</code></li></ol>",
    )
  })

  it("escapes <, > and & in text, so content can't become markup", () => {
    expect(markdownToTeamsHtml("a < b & c > d <script>alert(1)</script>")).toBe(
      "<p>a &lt; b &amp; c &gt; d &lt;script&gt;alert(1)&lt;/script&gt;</p>",
    )
  })

  it("keeps paragraphs and line breaks", () => {
    expect(markdownToTeamsHtml("Hi team,\nquick note.\n\nThanks")).toBe("<p>Hi team,<br>quick note.</p><p>Thanks</p>")
  })

  it("bolds and italicises, but leaves snake_case and lone stars alone", () => {
    expect(markdownToTeamsHtml("**bold**, *it* and _also_")).toBe("<p><b>bold</b>, <i>it</i> and <i>also</i></p>")
    expect(markdownToTeamsHtml("set MS365_EXTRA_SCOPES to 2 * 3")).toBe("<p>set MS365_EXTRA_SCOPES to 2 * 3</p>")
  })

  it("keeps an http(s) link and drops any other scheme to text", () => {
    expect(markdownToTeamsHtml("[docs](https://example.com/a?b=1&c=2)")).toBe(
      '<p><a href="https://example.com/a?b=1&amp;c=2">docs</a></p>',
    )
    expect(markdownToTeamsHtml("[x](javascript:alert(1))")).not.toContain("<a")
    expect(markdownToTeamsHtml('[x](https://e.com/"onmouseover=1)')).not.toContain('"onmouseover')
  })

  it("renders a heading as a bold line", () => {
    expect(markdownToTeamsHtml("## Setup\nsteps")).toBe("<p><b>Setup</b></p><p>steps</p>")
  })

  it("handles a long message without running out of stack", () => {
    const long = Array.from({ length: 20000 }, (_, k) => (k % 2 === 0 ? `line ${k}` : "")).join("\n")
    expect(markdownToTeamsHtml(long).startsWith("<p>line 0</p><p>line 2</p>")).toBe(true)
  })

  it("keeps bold and italic that wrap inline code or a link", () => {
    expect(markdownToTeamsHtml("**run `pnpm i` first**")).toBe("<p><b>run <code>pnpm i</code> first</b></p>")
    expect(markdownToTeamsHtml("*see [docs](https://e.com) now*")).toBe(
      '<p><i>see <a href="https://e.com/">docs</a> now</i></p>',
    )
  })

  it("never touches emphasis markers inside inline code", () => {
    expect(markdownToTeamsHtml("`**not bold**` and **bold**")).toBe("<p><code>**not bold**</code> and <b>bold</b></p>")
  })

  it("nests ***both*** properly, and leaves a crossed pair alone", () => {
    expect(markdownToTeamsHtml("***x***")).toBe("<p><b><i>x</i></b></p>")
    expect(markdownToTeamsHtml("**a *b** c*")).toBe("<p><b>a *b</b> c*</p>")
  })

  it("leaves an image as text rather than half-converting it", () => {
    expect(markdownToTeamsHtml("![logo](https://e.com/a.png)")).toBe("<p>![logo](https://e.com/a.png)</p>")
  })

  it("can't be tricked by a placeholder character in the text into substituting a placeholder", () => {
    expect(markdownToTeamsHtml("`<b>` then \uE0000\uE000")).toBe("<p><code>&lt;b&gt;</code> then 0</p>")
  })

  it("reads Windows line endings", () => {
    expect(markdownToTeamsHtml("- a\r\n- b\r\n\r\ndone")).toBe("<ul><li>a</li><li>b</li></ul><p>done</p>")
  })

  it("converts 100,000 lines in well under a second", () => {
    const lines = Array.from({ length: 100_000 }, (_, k) => (k % 2 === 0 ? `# heading ${k}` : ""))
    const started = performance.now()
    markdownToTeamsHtml(lines.join("\n"))
    expect(performance.now() - started).toBeLessThan(1000)
  })

  it("runs an unclosed code block to the end", () => {
    expect(markdownToTeamsHtml("```\nx < y")).toBe("<pre><code>x &lt; y</code></pre>")
  })
})

describe("messageFormatProblem", () => {
  it("refuses Markdown over the cap, and never refuses text", () => {
    const long = "x".repeat(MAX_MARKDOWN_CHARS + 1)
    expect(messageFormatProblem(long, "markdown")).toContain(
      "Split it into several messages, or send it with format: text.",
    )
    expect(messageFormatProblem(long, "text")).toBeUndefined()
    expect(messageFormatProblem(long, undefined)).toBeUndefined()
    expect(messageFormatProblem("x".repeat(MAX_MARKDOWN_CHARS), "markdown")).toBeUndefined()
  })
})

describe("messageBody", () => {
  it("sends text unchanged by default", () => {
    expect(messageBody("a < b\n*hi*")).toEqual({ contentType: "text", content: "a < b\n*hi*" })
  })

  it("sends markdown as HTML", () => {
    expect(messageBody("- a", "markdown")).toEqual({ contentType: "html", content: "<ul><li>a</li></ul>" })
  })
})
