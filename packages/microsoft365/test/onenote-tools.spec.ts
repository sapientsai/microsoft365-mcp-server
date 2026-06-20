import { Some } from "functype"
import { Left, Right } from "functype/either"
import { beforeEach, describe, expect, it, vi } from "vitest"

vi.mock("../src/client/graph-client", () => ({
  getGraphClient: vi.fn(),
}))

import { getGraphClient } from "../src/client/graph-client"
import {
  copyOnenotePage,
  createOnenoteNotebook,
  createOnenotePage,
  createOnenoteSection,
  deleteOnenotePage,
  listOnenoteNotebooks,
  updateOnenotePageContent,
} from "../src/tools/onenote-tools"

const mockClient = {
  listOnenoteNotebooks: vi.fn(),
  createOnenotePage: vi.fn(),
  updateOnenotePageContent: vi.fn(),
  createOnenoteSection: vi.fn(),
  createOnenoteNotebook: vi.fn(),
  copyOnenotePage: vi.fn(),
  deleteOnenotePage: vi.fn(),
}

beforeEach(() => {
  vi.clearAllMocks()
  vi.mocked(getGraphClient).mockReturnValue(Some(mockClient as never))
})

describe("onenote-tools", () => {
  describe("listOnenoteNotebooks (renamed read)", () => {
    it("should list notebooks via the renamed client method", async () => {
      mockClient.listOnenoteNotebooks.mockResolvedValue(Right({ value: [{ displayName: "Work" }] }))
      const result = await listOnenoteNotebooks({})
      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("Work")
      expect(mockClient.listOnenoteNotebooks).toHaveBeenCalled()
    })
  })

  describe("createOnenotePage", () => {
    it("should send an HTML document containing the title and return the page ID", async () => {
      mockClient.createOnenotePage.mockResolvedValue(Right({ id: "page-1" }))
      const result = await createOnenotePage({ section_id: "sec-1", title: "Notes", content: "<p>hello</p>" })
      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("page-1")
      const [sectionId, html] = mockClient.createOnenotePage.mock.calls[0] as [string, string]
      expect(sectionId).toBe("sec-1")
      expect(html).toContain("<title>Notes</title>")
      expect(html).toContain("<p>hello</p>")
    })

    it("should escape HTML-significant characters in the title", async () => {
      mockClient.createOnenotePage.mockResolvedValue(Right({ id: "page-2" }))
      await createOnenotePage({ section_id: "sec-1", title: "A & B <x>", content: "<p>x</p>" })
      const [, html] = mockClient.createOnenotePage.mock.calls[0] as [string, string]
      expect(html).toContain("<title>A &amp; B &lt;x&gt;</title>")
    })
  })

  describe("updateOnenotePageContent", () => {
    it("should default to append on body and wrap the command in an array", async () => {
      mockClient.updateOnenotePageContent.mockResolvedValue(Right({}))
      const result = await updateOnenotePageContent({ page_id: "page-1", content: "<p>more</p>" })
      expect(result.isRight()).toBe(true)
      expect(mockClient.updateOnenotePageContent).toHaveBeenCalledWith("page-1", [
        { target: "body", action: "append", content: "<p>more</p>" },
      ])
    })

    it("should pass through explicit action/target/position", async () => {
      mockClient.updateOnenotePageContent.mockResolvedValue(Right({}))
      await updateOnenotePageContent({
        page_id: "page-1",
        content: "<p>x</p>",
        action: "insert",
        target: "div:0",
        position: "after",
      })
      expect(mockClient.updateOnenotePageContent).toHaveBeenCalledWith("page-1", [
        { target: "div:0", action: "insert", content: "<p>x</p>", position: "after" },
      ])
    })
  })

  describe("createOnenoteSection", () => {
    it("should create a section and return its ID", async () => {
      mockClient.createOnenoteSection.mockResolvedValue(Right({ id: "sec-9" }))
      const result = await createOnenoteSection({ notebook_id: "nb-1", display_name: "Q2" })
      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("sec-9")
      expect(mockClient.createOnenoteSection).toHaveBeenCalledWith("nb-1", "Q2")
    })
  })

  describe("createOnenoteNotebook", () => {
    it("should create a notebook and return its ID", async () => {
      mockClient.createOnenoteNotebook.mockResolvedValue(Right({ id: "nb-9" }))
      const result = await createOnenoteNotebook({ display_name: "Research" })
      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("nb-9")
      expect(mockClient.createOnenoteNotebook).toHaveBeenCalledWith("Research")
    })
  })

  describe("copyOnenotePage", () => {
    it("should initiate a copy to the destination section", async () => {
      mockClient.copyOnenotePage.mockResolvedValue(Right({}))
      const result = await copyOnenotePage({ page_id: "page-1", section_id: "sec-2" })
      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("sec-2")
      expect(mockClient.copyOnenotePage).toHaveBeenCalledWith("page-1", "sec-2")
    })
  })

  describe("deleteOnenotePage", () => {
    it("should delete a page by ID", async () => {
      mockClient.deleteOnenotePage.mockResolvedValue(Right({}))
      const result = await deleteOnenotePage({ page_id: "page-1" })
      expect(result.isRight()).toBe(true)
      expect(result.value).toContain("deleted")
      expect(mockClient.deleteOnenotePage).toHaveBeenCalledWith("page-1")
    })

    it("should surface a UserError when the client fails", async () => {
      mockClient.deleteOnenotePage.mockResolvedValue(Left({ message: "Not found" }))
      const result = await deleteOnenotePage({ page_id: "missing" })
      expect(result.isLeft()).toBe(true)
      expect((result.value as Error).message).toContain("Failed to delete page")
    })
  })
})
