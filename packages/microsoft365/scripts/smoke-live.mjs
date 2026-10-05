#!/usr/bin/env node
// Drives the BUILT server against real Microsoft Graph, before a release.
//
// The unit suite replaces Graph with a fake that accepts any query, so it cannot see Graph's own
// rules. 1.2.9 shipped with list_messages unable to filter on importance ("InefficientFilter") while
// 299 tests passed. This script is the missing check: it starts dist/bin.js over stdio, speaks MCP to
// it, and calls the tools with the queries callers actually send.
//
// It needs a real sign-in, so it is not part of validate or CI. Run it before tagging:
//
//   pnpm build && pnpm --filter microsoft365-mcp-server smoke:live
//
// In the default interactive mode the server blocks at startup until the Microsoft sign-in in the
// browser completes. Credentials come from the repo-root .env (MS365_CLIENT_ID, MS365_TENANT_ID, ...)
// or the shell; MS365_ORG_MODE=true is needed for the meeting tools.
//
// Data-dependent checks SKIP rather than fail when the account has nothing to test with. Set
// SMOKE_MEETING_ID to page through that meeting's first transcript.
//
// Output is tool names, pass/fail/skip and counts only. Tool output can hold mail and file contents
// and lands wherever this is run (an agent transcript included), so it is never printed; nor is the
// environment. Server stderr is forwarded only for its setup and sign-in lines.

import { spawn } from "node:child_process"
import { existsSync } from "node:fs"
import { dirname, join } from "node:path"
import { createInterface } from "node:readline"
import { fileURLToPath } from "node:url"

const PACKAGE_ROOT = join(dirname(fileURLToPath(import.meta.url)), "..")
const REPO_ROOT = join(PACKAGE_ROOT, "..", "..")
const STARTUP_TIMEOUT_MS = 300_000 // covers a browser sign-in
const CALL_TIMEOUT_MS = 90_000

const FORWARDED_STDERR = /^\s*(\[(Setup|Server|Error|Auth|Fatal)\]|Authentication Required|Please visit|And enter code)/

const BIN = join(PACKAGE_ROOT, "dist", "bin.js")
if (!existsSync(BIN)) {
  console.error(`✗ ${BIN} is missing — run \`pnpm build\` first.`)
  process.exit(1)
}

// The server loads .env from its working directory, so it runs from the repo root.
const server = spawn(process.execPath, [BIN], {
  cwd: REPO_ROOT,
  env: { ...process.env, TRANSPORT_TYPE: "stdio" },
  stdio: ["pipe", "pipe", "pipe"],
})

// A write after the server has exited raises EPIPE; the exit handler below already reports it.
server.stdin.on("error", () => {})

createInterface({ input: server.stderr }).on("line", (line) => {
  if (FORWARDED_STDERR.test(line)) console.error(`  server: ${line.trim().slice(0, 200)}`)
})

const pending = new Map()
let nextId = 1

createInterface({ input: server.stdout }).on("line", (line) => {
  let message
  try {
    message = JSON.parse(line)
  } catch {
    return // not JSON-RPC; never echoed, since it could hold anything
  }
  // A request from the server (ping, roots/list) can carry the same id as one of ours; only a
  // response may resolve a pending call.
  if (message.method !== undefined) return
  const waiter = pending.get(message.id)
  if (waiter) {
    pending.delete(message.id)
    waiter(message)
  }
})

let exited = false
server.on("exit", (code) => {
  exited = true
  for (const waiter of pending.values()) waiter({ error: { message: `server exited (code ${code})` } })
  pending.clear()
})

const rpc = (method, params, timeoutMs) =>
  new Promise((resolve) => {
    if (exited) return resolve({ error: { message: "server is not running" } })
    const id = nextId++
    const timer = setTimeout(() => {
      pending.delete(id)
      resolve({ error: { message: `no response within ${timeoutMs / 1000}s` } })
    }, timeoutMs)
    pending.set(id, (message) => {
      clearTimeout(timer)
      resolve(message)
    })
    server.stdin.write(`${JSON.stringify({ jsonrpc: "2.0", id, method, params })}\n`)
  })

const firstLine = (text) =>
  String(text ?? "")
    .split("\n")[0]
    .slice(0, 200)

/** Calls a tool. Returns { ok, text } or { ok: false, error }. */
const call = async (name, args) => {
  const response = await rpc("tools/call", { name, arguments: args }, CALL_TIMEOUT_MS)
  if (response.error) return { ok: false, error: firstLine(response.error.message) }
  const text = (response.result?.content ?? []).map((part) => part.text ?? "").join("\n")
  if (response.result?.isError) return { ok: false, error: firstLine(text) }
  return { ok: true, text }
}

const results = []
const record = (status, name, detail) => {
  results.push(status)
  console.log(`${status.padEnd(4)}  ${name}${detail ? ` — ${detail}` : ""}`)
}

const messageLines = (text) => text.split("\n").filter((line) => line.startsWith("- **"))

// One list_messages case: Graph must accept the query, and every line must end with the Graph ID.
const listCase = async (label, args) => {
  const result = await call("list_messages", args)
  if (!result.ok) return record("FAIL", label, result.error)
  const lines = messageLines(result.text)
  const malformed = lines.filter((line) => !/\(ID: [^)]+\)$/.test(line))
  if (malformed.length > 0) return record("FAIL", label, `${malformed.length} line(s) do not end with "(ID: ...)"`)
  record("PASS", label, `${lines.length} message(s)`)
  return lines
}

const main = async () => {
  console.log("Starting the server (sign in in the browser if one opens)...")
  const init = await rpc(
    "initialize",
    { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "smoke-live", version: "1" } },
    STARTUP_TIMEOUT_MS,
  )
  if (init.error) {
    record("FAIL", "server start", `${firstLine(init.error.message)} — see the server lines above`)
    return
  }
  server.stdin.write(`${JSON.stringify({ jsonrpc: "2.0", method: "notifications/initialized" })}\n`)

  const me = await call("get_me", {})
  if (!me.ok) {
    record("FAIL", "get_me", me.error)
    return // nothing below can work without a signed-in user
  }
  record("PASS", "get_me")

  // Mail: the filters callers send. Each would fail with InefficientFilter if the sort property did
  // not lead the filter Graph receives.
  const weekAgo = new Date(Date.now() - 7 * 24 * 3600 * 1000).toISOString().replace(/\.\d+Z$/, "Z")
  await listCase("list_messages", { top: 3 })
  const high = await listCase("list_messages filter importance eq 'high'", { top: 3, filter: "importance eq 'high'" })
  await listCase("list_messages filter isRead eq false", { top: 3, filter: "isRead eq false" })
  await listCase("list_messages filter receivedDateTime ge (7 days ago)", {
    top: 3,
    filter: `receivedDateTime ge ${weekAgo}`,
  })
  await listCase("list_messages folder inbox, with previews", { top: 3, folder: "inbox", include_preview: true })
  await listCase("list_messages folder sentitems", { top: 3, folder: "sentitems" })
  await listCase("list_messages folder inbox, fetch_all_pages, date filter", {
    folder: "inbox",
    fetch_all_pages: true,
    filter: `receivedDateTime ge ${weekAgo}`,
  })
  // Leads with the sort property but adds an "or", so it is sent unprefixed. If Graph rejects this,
  // orderableFilter should prefix every filter instead.
  await listCase("list_messages filter receivedDateTime ... or importance", {
    top: 3,
    filter: `receivedDateTime ge ${weekAgo} or importance eq 'high'`,
  })

  // The prefix must exclude nothing. Drafts are where a missing receivedDateTime would show. The
  // prefix condition is sent on its own (it leads with receivedDateTime, so it goes through
  // unchanged), isolating exclusion from grouping; the counts must match the unfiltered call.
  const drafts = await call("list_messages", { folder: "drafts", fetch_all_pages: true })
  const draftsPrefixed = await call("list_messages", {
    folder: "drafts",
    fetch_all_pages: true,
    filter: "receivedDateTime ge 1900-01-01T00:00:00Z",
  })
  if (!drafts.ok || !draftsPrefixed.ok) {
    record("FAIL", "date prefix excludes no drafts", (drafts.ok ? draftsPrefixed : drafts).error)
  } else {
    const [plain, prefixed] = [messageLines(drafts.text).length, messageLines(draftsPrefixed.text).length]
    if (plain === 0) record("SKIP", "date prefix excludes no drafts", "no drafts")
    else if (plain === prefixed) record("PASS", "date prefix excludes no drafts", `${plain} draft(s) both ways`)
    else record("FAIL", "date prefix excludes no drafts", `${plain} without the prefix, ${prefixed} with it`)
  }

  if (!high || high.length === 0) record("SKIP", "[High importance] flag", "no high-importance mail")
  else if (high.every((line) => line.includes("[High importance]"))) record("PASS", "[High importance] flag")
  else record("FAIL", "[High importance] flag", "a high-importance message printed without the flag")

  const search = await call("search_messages", { query: "meeting", top: 3 })
  if (!search.ok) record("FAIL", "search_messages", search.error)
  else record("PASS", "search_messages", `${messageLines(search.text).length} message(s)`)

  // Chats: named and newest first. Graph documents the members expand, the lastMessagePreview expand
  // and the sort separately, never together, so Graph accepting the combination is part of the check.
  // A one-on-one or untitled group still labelled by its type means the members expand was dropped.
  const chats = await call("list_chats", { top: 5 })
  if (!chats.ok) {
    record("FAIL", "list_chats names and newest-first order", chats.error)
  } else {
    const lines = chats.text.split("\n").filter((line) => line.startsWith("- **"))
    const times = lines.flatMap((line) => line.match(/\(last message ([^ )]+)/)?.[1] ?? []).map((t) => Date.parse(t))
    const newestFirst = times.every((time, i) => i === 0 || times[i - 1] >= time)
    const unnamed = lines.filter((line) => /^- \*\*(group|oneOnOne)\*\*/.test(line)).length
    if (lines.length === 0) record("SKIP", "list_chats names and newest-first order", "no chats")
    else if (times.length === 0)
      record("FAIL", "list_chats names and newest-first order", "no line shows a last message")
    else if (times.some(Number.isNaN))
      record("FAIL", "list_chats names and newest-first order", "unparseable message time")
    else if (!newestFirst) record("FAIL", "list_chats names and newest-first order", "not sorted newest first")
    else if (unnamed > 0)
      record("FAIL", "list_chats names and newest-first order", `${unnamed} chat(s) named only by type`)
    else record("PASS", "list_chats names and newest-first order", `${lines.length} chat(s)`)
  }

  // Chats since: a time just ahead of now must return nothing, and a week back must return only chats
  // whose last message is newer than it.
  const future = new Date(Date.now() + 60_000).toISOString()
  const weekBack = new Date(Date.now() - 7 * 24 * 3600 * 1000).toISOString()
  const noChats = await call("list_chats", { since: future })
  const recentChats = await call("list_chats", { since: weekBack })
  if (!noChats.ok || !recentChats.ok) {
    record("FAIL", "list_chats since", (noChats.ok ? recentChats : noChats).error)
  } else {
    const strays = noChats.text.split("\n").filter((line) => line.startsWith("- **")).length
    const recent = recentChats.text.split("\n").filter((line) => line.startsWith("- **"))
    const older = recent.filter((line) => {
      const time = Date.parse(line.match(/\(last message ([^ )]+)/)?.[1] ?? "")
      return Number.isNaN(time) || time <= Date.parse(weekBack)
    }).length
    if (strays > 0) record("FAIL", "list_chats since", `a future since returned ${strays} chat(s)`)
    else if (older > 0) record("FAIL", "list_chats since", `${older} chat(s) not newer than since`)
    else record("PASS", "list_chats since", `${recent.length} chat(s) active in the last 7 days`)
  }

  // Chat messages: every header line ends with the ID, no text line carries HTML, and a since just
  // ahead of now returns nothing. Graph ignores a lastModifiedDateTime filter that is not paired with
  // the matching $orderby, so messages coming back here would mean the filter was dropped.
  const chatId = chats.ok ? chats.text.match(/, ID: ([^)]+)\)$/m)?.[1] : undefined
  if (!chatId) {
    record("SKIP", "list_chat_messages format and since", "no chat to read")
  } else {
    const messages = await call("list_chat_messages", { chat_id: chatId, top: 5 })
    const noMessages = await call("list_chat_messages", { chat_id: chatId, since: future })
    if (!messages.ok || !noMessages.ok) {
      record("FAIL", "list_chat_messages format and since", (messages.ok ? noMessages : messages).error)
    } else {
      const headers = messages.text.split("\n").filter((line) => line.startsWith("- **"))
      const texts = messages.text.split("\n").filter((line) => line.startsWith("  > "))
      const badHeaders = headers.filter((line) => !/\(ID: [^)]+\)$/.test(line)).length
      const htmlTexts = texts.filter((line) => /<\/?(div|p|span|at|br|img|attachment)\b/i.test(line)).length
      const leaked = noMessages.text.split("\n").filter((line) => line.startsWith("- **")).length
      if (badHeaders > 0)
        record("FAIL", "list_chat_messages format and since", `${badHeaders} line(s) not ending in the ID`)
      else if (htmlTexts > 0)
        record("FAIL", "list_chat_messages format and since", `${htmlTexts} text line(s) with HTML`)
      else if (/could not be resolved/.test(messages.text))
        record("FAIL", "list_chat_messages format and since", "the signed-in user was not resolved")
      else if (leaked > 0)
        record("FAIL", "list_chat_messages format and since", `a future since returned ${leaked}: filter ignored`)
      else
        record("PASS", "list_chat_messages format and since", `${headers.length} message(s), ${texts.length} with text`)
    }
  }

  // Search: top must be honoured, not just accepted. A larger top returning more proves the cap.
  const searchTwo = await call("search_files", { query: "a", top: 2 })
  const searchFive = await call("search_files", { query: "a", top: 5 })
  if (!searchTwo.ok || !searchFive.ok) {
    record("FAIL", "search_files top", (searchTwo.ok ? searchFive : searchTwo).error)
  } else {
    const [two, five] = [messageLines(searchTwo.text).length, messageLines(searchFive.text).length]
    if (two > 2) record("FAIL", "search_files top", `top 2 returned ${two} results`)
    else if (five <= 2) record("SKIP", "search_files top", "too few matches to show the cap")
    else record("PASS", "search_files top", `top 2 → ${two}, top 5 → ${five}`)
  }

  // Search folders: a folder that list_drive_items shows as non-empty must not read "0 items" in search.
  const root = await call("list_drive_items", {})
  const folder = root.ok
    ? [...root.text.matchAll(/^- \*\*([^*]+)\*\* \(ID: ([^)]+)\) - Folder \((\d+) items\)/gm)].find(
        (match) => Number(match[3]) > 0,
      )
    : undefined
  if (!root.ok) {
    record("FAIL", "search folder line", root.error)
  } else if (!folder) {
    record("SKIP", "search folder line", "no non-empty folder at the OneDrive root")
  } else {
    const found = await call("search_files", { query: folder[1], top: 25 })
    const line = found.ok
      ? found.text.split("\n").find((candidate) => candidate.includes(`(ID: ${folder[2]})`))
      : undefined
    if (!found.ok) record("FAIL", "search folder line", found.error)
    else if (!line) record("SKIP", "search folder line", "search did not return the folder")
    // Matched right after the ID, so a folder named "Folder (2020)" or a path holding "0 B" cannot trip it.
    else if (/\) - Folder \(0 items\)|\) - Folder (\(\d+ items\) )?\(0 B\)/.test(line))
      record("FAIL", "search folder line", "still shows a zero count or size")
    else record("PASS", "search folder line", `${folder[3]}-item folder shown without a zero count`)
  }

  // Events: the body must arrive as text, not the Teams invite's HTML. A recent window, because
  // list_events sorts oldest first; and an empty body proves nothing, so the first non-empty one counts.
  const day = 24 * 3600 * 1000
  const window = await call("list_calendar_view", {
    start_date_time: new Date(Date.now() - 14 * day).toISOString(),
    end_date_time: new Date(Date.now() + 14 * day).toISOString(),
  })
  if (!window.ok) {
    record("FAIL", "get_event text body", window.error)
  } else {
    const eventIds = [...window.text.matchAll(/\(ID: ([^)]+)\)$/gm)].map((match) => match[1]).slice(0, 8)
    let outcome = ["SKIP", eventIds.length === 0 ? "no events in the last or next 14 days" : "no event with a body"]
    for (const eventId of eventIds) {
      const event = await call("get_event", { event_id: eventId })
      if (!event.ok) {
        outcome = ["FAIL", event.error]
        break
      }
      const body = (event.text.split("## Body")[1] ?? "").trim()
      if (body.length === 0) continue
      outcome = /<(html|head|body|div|p|br|span|table|meta)\b/i.test(body)
        ? ["FAIL", "body has HTML tags"]
        : ["PASS", `${body.length} chars of text`]
      break
    }
    record(outcome[0], "get_event text body", outcome[1])
  }

  // Files: a .log OneDrive stored with an unknown type must come back as text.
  const files = await call("search_files", { query: "log" })
  if (!files.ok) {
    record("FAIL", "search_files", files.error)
  } else {
    record("PASS", "search_files")
    // The type after the ID is the MIME type (octet-stream for a .log), so only the name and ID are matched.
    const log = files.text.match(/^- \*\*[^*]+\.log\*\* \(ID: ([^)]+)\)/m)
    if (!log) {
      record("SKIP", "download_file / read_document on a .log", "no .log file found")
    } else {
      const download = await call("download_file", { item_id: log[1] })
      if (!download.ok) record("FAIL", "download_file .log", download.error)
      else if (download.text.includes("## Content")) record("PASS", "download_file .log", "content shown")
      else record("SKIP", "download_file .log", "over the 100 KB inline limit, or not text")
      const read = await call("read_document", { path: `/me/drive/items/${log[1]}/content`, max_chars: 2000 })
      record(read.ok ? "PASS" : "FAIL", "read_document .log", read.ok ? "" : read.error)
    }
  }

  // Transcripts: page through until the last part, following each marker's offset.
  const meetingId = process.env.SMOKE_MEETING_ID
  if (!meetingId) {
    record("SKIP", "transcript paging", "set SMOKE_MEETING_ID to test it")
    return
  }
  const transcripts = await call("list_meeting_transcripts", { meeting_id: meetingId })
  const transcriptId = transcripts.ok ? transcripts.text.match(/transcript_id: `([^`]+)`/)?.[1] : undefined
  if (!transcripts.ok) return record("FAIL", "list_meeting_transcripts", transcripts.error)
  if (!transcriptId) return record("SKIP", "transcript paging", "the meeting has no transcript")

  let offset = 0
  for (let part = 1; part <= 20; part++) {
    const page = await call("get_meeting_transcript", {
      meeting_id: meetingId,
      transcript_id: transcriptId,
      max_chars: 20_000,
      offset,
    })
    if (!page.ok) return record("FAIL", "transcript paging", `part ${part}: ${page.error}`)
    const next = page.text.match(/call again with offset: (\d+)\]$/)
    if (!next) return record("PASS", "transcript paging", `${part} part(s)`)
    if (Number(next[1]) <= offset) return record("FAIL", "transcript paging", "the offset did not advance")
    offset = Number(next[1])
  }
  record("FAIL", "transcript paging", "no end marker after 20 parts")
}

try {
  await main()
} finally {
  server.kill() // an orphaned server could keep the sign-in redirect port and break the next run
}

const failed = results.filter((status) => status === "FAIL").length
const skipped = results.filter((status) => status === "SKIP").length
console.log(
  `\n${failed === 0 ? "✔" : "✗"} smoke:live — ${results.length - failed - skipped} passed, ${failed} failed, ${skipped} skipped`,
)
process.exit(failed === 0 ? 0 : 1)
