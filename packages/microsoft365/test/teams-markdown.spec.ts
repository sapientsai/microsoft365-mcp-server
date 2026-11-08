import { describe, expect, it } from "vitest"

import { markdownToTeamsHtml, messageBody } from "../src/utils/teams-markdown"

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

  it("runs an unclosed code block to the end", () => {
    expect(markdownToTeamsHtml("```\nx < y")).toBe("<pre><code>x &lt; y</code></pre>")
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
