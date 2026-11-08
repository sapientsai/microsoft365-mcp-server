import { Right } from "functype/either"
import { describe, expect, it, vi } from "vitest"

import { createGraphRequest } from "../src/graph-request"
import type { GraphApiError } from "../src/types"
import { fetchAllPages, nextLinkProblem } from "../src/utils/pagination"

const page = (value: ReadonlyArray<number>, next?: string) =>
  Right<GraphApiError, { value: number[]; "@odata.nextLink"?: string }>({
    value: [...value],
    ...(next === undefined ? {} : { "@odata.nextLink": next }),
  })

describe("nextLinkProblem", () => {
  it.each([
    "https://graph.microsoft.com/v1.0/me/messages?$skiptoken=abc",
    "https://graph.microsoft.com/beta/me/chats?$skiptoken=abc",
    "https://GRAPH.microsoft.com:443/v1.0/groups?$skiptoken=x",
    "https://graph.microsoft.com/v1.0/me/mailFolders('AAMkAD%3D')/messages?$skip=10",
    "https://graph.microsoft.com/v1.0/me/drive/root/search(q='foo%20bar')?$skiptoken=y",
    "https://graph.microsoft.com/v1.0/users/delta?$deltatoken=abc",
  ])("allows Graph's own link %s", (link) => {
    expect(nextLinkProblem(link)).toBeUndefined()
  })

  it.each([
    ["another host", "https://attacker.example/v1.0/me/messages", "points at https://attacker.example"],
    ["a look-alike host", "https://graph.microsoft.com.attacker.example/v1.0/me", "points at"],
    ["plain http", "http://graph.microsoft.com/v1.0/me/messages", "points at http://graph.microsoft.com"],
    ["another port", "https://graph.microsoft.com:8443/v1.0/me", "points at"],
    ["credentials", "https://user:pass@graph.microsoft.com/v1.0/me/messages", "carries credentials"],
    ["a fragment", "https://graph.microsoft.com/v1.0/me/messages#x", "carries a fragment"],
    ["no version", "https://graph.microsoft.com/me/messages", "isn't under /v1.0/ or /beta/"],
    ["a relative link", "/v1.0/me/messages", "isn't a URL"],
  ])("refuses %s", (_name, link, reason) => {
    expect(nextLinkProblem(link)).toContain(reason)
  })
})

describe("fetchAllPages", () => {
  it("follows Graph's next links to the end", async () => {
    const fetchPage = vi
      .fn()
      .mockResolvedValueOnce(page([1, 2], "https://graph.microsoft.com/v1.0/me/messages?$skiptoken=2"))
      .mockResolvedValueOnce(page([3]))
    const result = await fetchAllPages(fetchPage, "https://graph.microsoft.com/v1.0/me/messages")
    expect(result.value).toEqual([1, 2, 3])
    expect(fetchPage).toHaveBeenCalledTimes(2)
  })

  it.each([
    ["no link", undefined],
    ["an empty link", ""],
  ])("ends paging on %s", async (_name, next) => {
    const fetchPage = vi.fn().mockResolvedValueOnce(page([1, 2], next))
    const result = await fetchAllPages(fetchPage, "https://graph.microsoft.com/v1.0/me/messages")
    expect(result.value).toEqual([1, 2])
    expect(fetchPage).toHaveBeenCalledTimes(1)
  })

  it("fetches the next page by its parsed URL", async () => {
    const fetchPage = vi
      .fn()
      .mockResolvedValueOnce(page([1], "https://GRAPH.microsoft.com:443/v1.0/me/messages?$skiptoken=2"))
      .mockResolvedValueOnce(page([2]))
    await fetchAllPages(fetchPage, "https://graph.microsoft.com/v1.0/me/messages")
    expect(fetchPage).toHaveBeenLastCalledWith("https://graph.microsoft.com/v1.0/me/messages?$skiptoken=2")
  })

  it("stops with an error, and requests nothing more, when a next link points elsewhere", async () => {
    const fetchPage = vi.fn().mockResolvedValueOnce(page([1], "https://attacker.example/v1.0/me/messages"))
    const result = await fetchAllPages(fetchPage, "https://graph.microsoft.com/v1.0/me/messages")
    expect(result.isLeft()).toBe(true)
    expect((result.value as GraphApiError).message).toContain("Stopped paging after 1 page(s)")
    expect(fetchPage).toHaveBeenCalledTimes(1)
  })
})

describe("requestPaginated", () => {
  it("never sends the token to a host a next link names", async () => {
    const fetchSpy = vi.fn((url: string) =>
      Promise.resolve(
        new Response(
          JSON.stringify(
            url.startsWith("https://graph.microsoft.com/")
              ? { value: [1], "@odata.nextLink": "https://user:pw@attacker.example/v1.0/me/messages" }
              : { value: [2] },
          ),
          { status: 200 },
        ),
      ),
    )
    const { requestPaginated } = createGraphRequest(
      { getAccessToken: () => Promise.resolve(Right("T")) },
      { fetchImpl: fetchSpy as unknown as typeof fetch },
    )
    const result = await requestPaginated<number>("/me/messages")
    expect(result.isLeft()).toBe(true)
    expect(fetchSpy).toHaveBeenCalledTimes(1)
    expect(fetchSpy.mock.calls.every(([url]) => String(url).startsWith("https://graph.microsoft.com/"))).toBe(true)
  })
})
