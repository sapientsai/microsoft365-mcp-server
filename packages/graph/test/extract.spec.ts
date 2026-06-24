import ExcelJS from "exceljs"
import { describe, expect, it } from "vitest"

import { EXTRACTABLE_TYPES, extractTextFromBuffer, isTextContent, resolveContentType } from "../src/extract/extract"

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

describe("isTextContent / resolveContentType", () => {
  it("recognizes text-ish content types", () => {
    expect(isTextContent("text/plain")).toBe(true)
    expect(isTextContent("application/json")).toBe(true)
    expect(isTextContent("application/vnd.oasis+xml")).toBe(true)
    expect(isTextContent("application/pdf")).toBe(false)
  })

  it("resolves octet-stream by extension", () => {
    expect(resolveContentType("application/octet-stream", "report.pdf")).toBe("application/pdf")
    expect(resolveContentType("", "data.xlsx")).toBe(XLSX)
    expect(resolveContentType("text/csv", "x.bin")).toBe("text/csv")
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

  it("rejects an unsupported content type", async () => {
    const result = await extractTextFromBuffer(Buffer.from("x"), "image/png", "p.png")
    expect(result.isLeft()).toBe(true)
    expect((result.value as { message: string }).message).toContain("Unsupported content type")
  })

  it("exposes the extractable types", () => {
    expect(EXTRACTABLE_TYPES).toContain("application/pdf")
    expect(EXTRACTABLE_TYPES).toContain(DOCX)
    expect(EXTRACTABLE_TYPES).toContain(XLSX)
  })
})
