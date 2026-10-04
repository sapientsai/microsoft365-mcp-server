import { describe, expect, it } from "vitest"

import { isTextMimeType, isUnknownMimeType, looksLikeText } from "../src/utils/text-detect"

const utf8 = (s: string): Uint8Array => new TextEncoder().encode(s)

// The first bytes of a real PNG. The IHDR chunk length (00 00 00 0d) puts NULs at offset 8.
const PNG_HEADER = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48])

describe("isTextMimeType", () => {
  it("accepts text/* and the text application types, ignoring parameters and case", () => {
    expect(isTextMimeType("text/plain")).toBe(true)
    expect(isTextMimeType("Text/CSV; charset=utf-8")).toBe(true)
    expect(isTextMimeType("application/json")).toBe(true)
    expect(isTextMimeType("application/javascript")).toBe(true)
    expect(isTextMimeType("application/vnd.oasis+xml")).toBe(true)
    expect(isTextMimeType("application/ld+json")).toBe(true)
  })

  it("rejects binary types, octet-stream and a missing type", () => {
    expect(isTextMimeType("application/pdf")).toBe(false)
    expect(isTextMimeType("image/png")).toBe(false)
    expect(isTextMimeType("application/octet-stream")).toBe(false)
    expect(isTextMimeType(undefined)).toBe(false)
  })
})

describe("isUnknownMimeType", () => {
  it("treats octet-stream and a missing type as unknown", () => {
    expect(isUnknownMimeType("application/octet-stream")).toBe(true)
    expect(isUnknownMimeType("application/octet-stream; foo=bar")).toBe(true)
    expect(isUnknownMimeType("")).toBe(true)
    expect(isUnknownMimeType(undefined)).toBe(true)
    expect(isUnknownMimeType("text/plain")).toBe(false)
  })
})

describe("looksLikeText", () => {
  it("accepts a .log and a .jsonl", () => {
    expect(looksLikeText(utf8("2026-10-03T10:00:00Z INFO started\n"), "app.log")).toBe(true)
    expect(looksLikeText(utf8('{"a":1}\n{"a":2}\n'), "events.jsonl")).toBe(true)
  })

  it("rejects a binary file renamed to .log", () => {
    expect(looksLikeText(PNG_HEADER, "image.log")).toBe(false)
  })

  it("accepts UTF-8 under an extension it does not know", () => {
    expect(looksLikeText(utf8("Grüße — café"), "notes.unknownext")).toBe(true)
  })

  it("rejects invalid UTF-8 under an extension it does not know", () => {
    expect(looksLikeText(new Uint8Array([0x48, 0x69, 0xff, 0xfe, 0x21]), "blob.bin")).toBe(false)
  })

  // Windows-1252 CSV exports are common and not valid UTF-8; a text extension plus no NULs is enough.
  it("accepts legacy-encoded bytes under a text extension", () => {
    expect(looksLikeText(new Uint8Array([0x63, 0x61, 0x66, 0xe9, 0x2c, 0x31]), "export.csv")).toBe(true)
  })

  it("does not reject UTF-8 because the 4 KB sample cuts a character in half", () => {
    const bytes = utf8(`${"a".repeat(4095)}é`) // é's two bytes straddle the 4096 boundary
    expect(looksLikeText(bytes, "data.unknownext")).toBe(true)
  })

  it("judges only the first 4 KB", () => {
    const bytes = new Uint8Array([...utf8("a".repeat(5000)), 0x00])
    expect(looksLikeText(bytes, "late-nul.bin")).toBe(true)
  })

  it("treats an empty file as text", () => {
    expect(looksLikeText(new Uint8Array(), "empty.bin")).toBe(true)
  })
})
