import ExcelJS from "exceljs"
import { describe, expect, it } from "vitest"

import { EXTRACTABLE_TYPES, extractTextFromBuffer, resolveContentType } from "../src/extract"

const DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
const XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

const xlsxBuffer = async (sheets: Record<string, string[][]>): Promise<Buffer> => {
  const wb = new ExcelJS.Workbook()
  for (const [name, rows] of Object.entries(sheets)) {
    const ws = wb.addWorksheet(name)
    for (const row of rows) ws.addRow(row)
  }
  return Buffer.from(await wb.xlsx.writeBuffer())
}

describe("resolveContentType", () => {
  it("resolves octet-stream to a document type by extension", () => {
    expect(resolveContentType("application/octet-stream", "report.pdf")).toBe("application/pdf")
    expect(resolveContentType("", "data.xlsx")).toBe(XLSX)
    expect(resolveContentType("text/csv", "x.bin")).toBe("text/csv")
  })

  // A text extension proves nothing about the bytes, so it is left unknown for the sniff to decide.
  it("leaves octet-stream unknown for a text extension", () => {
    expect(resolveContentType("application/octet-stream", "app.log")).toBe("application/octet-stream")
  })
})

describe("extractTextFromBuffer", () => {
  it("returns plain text / json / csv directly", async () => {
    expect((await extractTextFromBuffer(Buffer.from("Hello"), "text/plain", "n.txt")).value).toBe("Hello")
    const json = JSON.stringify({ a: 1 })
    expect((await extractTextFromBuffer(Buffer.from(json), "application/json", "d.json")).value).toBe(json)
    expect((await extractTextFromBuffer(Buffer.from("a,b\n1,2"), "text/csv", "d.csv")).value).toBe("a,b\n1,2")
  })

  it("extracts a single-sheet XLSX to CSV", async () => {
    const buf = await xlsxBuffer({
      Sheet1: [
        ["name", "age"],
        ["Alice", "30"],
      ],
    })
    const result = await extractTextFromBuffer(buf, XLSX, "data.xlsx")
    expect(result.isRight()).toBe(true)
    expect(result.value as string).toContain("name,age")
    expect(result.value as string).toContain("Alice,30")
  })

  it("labels sheets for a multi-sheet XLSX", async () => {
    const buf = await xlsxBuffer({ First: [["x"]], Second: [["y"]] })
    const text = (await extractTextFromBuffer(buf, XLSX, "data.xlsx")).value as string
    expect(text).toContain("[Sheet: First]")
    expect(text).toContain("[Sheet: Second]")
  })

  it("resolves an octet-stream XLSX by filename extension", async () => {
    const buf = await xlsxBuffer({ S: [["v"]] })
    expect((await extractTextFromBuffer(buf, "application/octet-stream", "x.xlsx")).isRight()).toBe(true)
  })

  it("returns a parse error for a corrupt DOCX", async () => {
    const result = await extractTextFromBuffer(Buffer.from("not a docx"), DOCX, "broken.docx")
    expect(result.isLeft()).toBe(true)
    expect((result.value as { type: string }).type).toBe("parse")
  })

  // Seen 2026-10-03: OneDrive stored a .log uploaded as text/plain as application/octet-stream,
  // and read_document refused it.
  it("reads an octet-stream .log and .jsonl by sniffing the bytes", async () => {
    const log = await extractTextFromBuffer(Buffer.from("INFO started\n"), "application/octet-stream", "app.log")
    expect(log.value).toBe("INFO started\n")
    const jsonl = await extractTextFromBuffer(Buffer.from('{"a":1}\n'), "application/octet-stream", "e.jsonl")
    expect(jsonl.value).toBe('{"a":1}\n')
  })

  it("refuses a binary file renamed to .log", async () => {
    const png = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d])
    const result = await extractTextFromBuffer(png, "application/octet-stream", "image.log")
    expect(result.isLeft()).toBe(true)
    expect((result.value as { type: string }).type).toBe("unsupported")
  })

  it("rejects an unsupported content type", async () => {
    const result = await extractTextFromBuffer(Buffer.from("x"), "image/png", "p.png")
    expect(result.isLeft()).toBe(true)
    expect((result.value as { message: string }).message).toContain("Unsupported content type")
  })

  // The two ExtractError arms are load-bearing for callers: "parse" means the file is broken,
  // "unsupported" means this package does not handle the format and download_file is the fallback.
  it("distinguishes unsupported from parse in the error type", async () => {
    const unsupported = await extractTextFromBuffer(Buffer.from("x"), "image/png", "p.png")
    expect((unsupported.value as { type: string }).type).toBe("unsupported")

    const corrupt = await extractTextFromBuffer(Buffer.from("not a docx"), DOCX, "broken.docx")
    expect((corrupt.value as { type: string }).type).toBe("parse")
  })

  it("exposes the extractable types", () => {
    expect(EXTRACTABLE_TYPES).toContain("application/pdf")
    expect(EXTRACTABLE_TYPES).toContain(DOCX)
    expect(EXTRACTABLE_TYPES).toContain(XLSX)
  })
})
