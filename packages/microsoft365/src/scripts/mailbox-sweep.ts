#!/usr/bin/env node
// Rule-driven mailbox clean-up.
//
// Reads every message in one folder, classifies each against a rules file, writes a
// per-sender report, and — only with --apply — moves the "delete" set to a destination
// folder (normally Deleted Items, so nothing is unrecoverable). The point of a script
// rather than an agent driving tool calls: the rules are reviewable before anything
// moves, the rehearsal report is the approval artefact, and a re-run is deterministic.
//
//   node dist/mailbox-sweep.js --rules rules.json [--out ./sweep-out] [--apply] [--pace [ms]]
//
// Auth is the server's own (same MS365_* environment, same token cache), so a cached
// sign-in is reused and no separate login flow exists.
//
// Classification, in order — the first rule that matches decides:
//   1. keep     sender/domain on the keep list, subject has a record word, has an
//               attachment, or is flagged. Keep always wins.
//   2. delete   sender on the explicit delete list.
//   3. delete   sender address looks like a marketing address AND its newest message
//               carries a List-Unsubscribe header (bulk mail; people and receipts
//               never have one) AND the subject has no record word.
//   4. keep     everything else (reported as "unmatched" so the rules can be tuned).

import { appendFileSync, mkdirSync, readFileSync, writeFileSync } from "node:fs"
import { join } from "node:path"

import { getAccessToken, initializeAuth } from "../auth"
import { getGraphClient, initializeGraphClient } from "../client/graph-client"
import { runMoveBatches } from "../mail/batch-move"
import { resolveMailboxScope } from "../mail/mailbox"
import type { AuthConfig, GraphBatchResponse, GraphMessage } from "../types"

type Rules = {
  readonly mailbox?: string
  readonly folder: string
  readonly destination: string
  readonly keep: {
    readonly addresses: ReadonlyArray<string>
    readonly domains: ReadonlyArray<string>
    readonly subjectWords: ReadonlyArray<string>
    readonly attachments: boolean
    readonly flagged: boolean
  }
  readonly delete: {
    readonly addresses: ReadonlyArray<string>
    readonly domains: ReadonlyArray<string>
    readonly marketingSubdomains: ReadonlyArray<string>
    readonly marketingLocalParts: ReadonlyArray<string>
    readonly requireUnsubscribeHeader: boolean
  }
  readonly ask: { readonly addresses: ReadonlyArray<string>; readonly domains: ReadonlyArray<string> }
}

type SweepMessage = GraphMessage & { readonly flag?: { readonly flagStatus?: string } }

type Decision =
  { readonly action: "keep"; readonly reason: string } | { readonly action: "delete"; readonly reason: string }

const FIELDS = ["id", "subject", "from", "receivedDateTime", "isRead", "hasAttachments", "flag"]

const arg = (name: string): string | undefined => {
  const i = process.argv.indexOf(name)
  return i >= 0 ? process.argv[i + 1] : undefined
}
const flag = (name: string): boolean => process.argv.includes(name)

const lower = (s: string | undefined): string => (s ?? "").trim().toLowerCase()
const addressOf = (m: GraphMessage): string => lower(m.from?.emailAddress.address)
const domainOf = (address: string): string => address.slice(address.lastIndexOf("@") + 1)
const localOf = (address: string): string => address.slice(0, address.lastIndexOf("@"))

// "example.com" matches example.com and any subdomain of it.
const inDomains = (address: string, domains: ReadonlyArray<string>): boolean => {
  const d = domainOf(address)
  return domains.some((rule) => d === rule.toLowerCase() || d.endsWith(`.${rule.toLowerCase()}`))
}
const inAddresses = (address: string, list: ReadonlyArray<string>): boolean =>
  list.some((a) => a.toLowerCase() === address)

const wordPattern = (words: ReadonlyArray<string>): RegExp =>
  new RegExp(`\\b(${words.map((w) => w.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")).join("|")})\\b`, "i")

const looksMarketing = (address: string, rules: Rules["delete"]): boolean => {
  const labels = domainOf(address).split(".")
  const sub = labels.length > 2 ? labels.slice(0, -2) : []
  const local = localOf(address).replace(/[^a-z]/g, "")
  return (
    sub.some((s) => rules.marketingSubdomains.includes(s)) ||
    rules.marketingLocalParts.some((p) => local === p || local.startsWith(p))
  )
}

const classify = (
  m: SweepMessage,
  rules: Rules,
  recordWord: RegExp,
  unsubscribeBySender: ReadonlyMap<string, boolean>,
): Decision => {
  const address = addressOf(m)
  const subject = m.subject ?? ""
  if (address === "") return { action: "keep", reason: "keep:no-sender" }
  if (inAddresses(address, rules.keep.addresses)) return { action: "keep", reason: "keep:address" }
  if (inDomains(address, rules.keep.domains)) return { action: "keep", reason: "keep:domain" }
  if (recordWord.test(subject)) return { action: "keep", reason: "keep:record-word" }
  if (rules.keep.attachments && m.hasAttachments) return { action: "keep", reason: "keep:attachment" }
  if (rules.keep.flagged && m.flag?.flagStatus === "flagged") return { action: "keep", reason: "keep:flagged" }
  if (inAddresses(address, rules.ask.addresses) || inDomains(address, rules.ask.domains))
    return { action: "keep", reason: "keep:ask" }
  if (inAddresses(address, rules.delete.addresses)) return { action: "delete", reason: "delete:address" }
  if (inDomains(address, rules.delete.domains)) return { action: "delete", reason: "delete:domain" }
  if (looksMarketing(address, rules.delete)) {
    const unsub = unsubscribeBySender.get(address)
    if (!rules.delete.requireUnsubscribeHeader || unsub === true)
      return { action: "delete", reason: "delete:marketing" }
    return { action: "keep", reason: unsub === false ? "keep:no-unsubscribe" : "keep:unsubscribe-unknown" }
  }
  return { action: "keep", reason: "keep:unmatched" }
}

const csvCell = (v: string | number): string => {
  const s = String(v)
  return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s
}
const csv = (rows: ReadonlyArray<ReadonlyArray<string | number>>): string =>
  rows.map((r) => r.map(csvCell).join(",")).join("\n")

const authConfig = (): AuthConfig => ({
  mode: "interactive",
  tenantId: process.env.MS365_TENANT_ID ?? "common",
  clientId: process.env.MS365_CLIENT_ID ?? "",
  redirectUri: process.env.MS365_REDIRECT_URI,
})

const fail = (message: string): never => {
  console.error(`[sweep] ${message}`)
  process.exit(1)
}

const main = async (): Promise<void> => {
  const rulesPath = arg("--rules") ?? fail("--rules <file> is required")
  const outDir = arg("--out") ?? "./sweep-out"
  const apply = flag("--apply")
  // --pace with no number means the default; absent means no pacing.
  const paceArg = arg("--pace")
  const paceMs = !flag("--pace") ? 0 : paceArg !== undefined && /^\d+$/.test(paceArg) ? Number(paceArg) : 750
  const rules = JSON.parse(readFileSync(rulesPath, "utf-8")) as Rules
  mkdirSync(outDir, { recursive: true })

  // A rules file without a mailbox would run against the signed-in user's own mail.
  // That is never an accident worth allowing silently.
  if (!rules.mailbox && !flag("--own-mailbox"))
    fail("rules.mailbox is not set; pass --own-mailbox to sweep your own mailbox")
  if (!rules.folder || !rules.destination) fail("rules.folder and rules.destination are required")

  const auth = await initializeAuth(authConfig())
  if (auth.isLeft()) fail(`auth failed: ${(auth.value as { message: string }).message}`)
  initializeGraphClient({ getAccessToken })
  const client = getGraphClient().orThrow()

  const scope = resolveMailboxScope(rules.mailbox)
  if (scope.isLeft()) fail((scope.value as { message: string }).message)
  const { prefix } = scope.orThrow()

  console.error(`[sweep] reading ${rules.folder} in ${rules.mailbox ?? "own mailbox"} …`)
  const fetched = await client.requestPaginated<SweepMessage>(`${prefix}/mailFolders/${rules.folder}/messages`, {
    odataParams: { $select: FIELDS, $top: 999 },
  })
  if (fetched.isLeft()) fail(`read failed: ${(fetched.value as { message: string }).message}`)
  const messages = fetched.orThrow()
  console.error(`[sweep] ${messages.length} messages`)

  // Newest message per sender, for the List-Unsubscribe probe.
  const newestBySender = new Map<string, SweepMessage>()
  messages.forEach((m) => {
    const a = addressOf(m)
    const cur = newestBySender.get(a)
    if (!cur || (m.receivedDateTime ?? "") > (cur.receivedDateTime ?? "")) newestBySender.set(a, m)
  })

  const recordWord = wordPattern(rules.keep.subjectWords)
  const probeTargets = [...newestBySender.entries()].filter(
    ([a]) =>
      a !== "" &&
      !inAddresses(a, rules.keep.addresses) &&
      !inDomains(a, rules.keep.domains) &&
      !inAddresses(a, rules.delete.addresses) &&
      !inDomains(a, rules.delete.domains) &&
      looksMarketing(a, rules.delete),
  )
  console.error(`[sweep] probing List-Unsubscribe on ${probeTargets.length} marketing-looking senders …`)

  const unsubscribeBySender = new Map<string, boolean>()
  const chunks = Array.from({ length: Math.ceil(probeTargets.length / 20) }, (_, i) =>
    probeTargets.slice(i * 20, i * 20 + 20),
  )
  await chunks.reduce<Promise<void>>(async (acc, chunk) => {
    await acc
    const requests = chunk.map(([a, m]) => ({
      id: a,
      method: "GET",
      url: `${prefix}/messages/${m.id}?$select=internetMessageHeaders`,
    }))
    const res = await client.request<GraphBatchResponse>("POST", "/$batch", { body: { requests } })
    if (res.isLeft()) {
      console.error(`[sweep] probe batch failed: ${(res.value as { message: string }).message}`)
      return
    }
    const batch = res.value as GraphBatchResponse
    batch.responses.forEach((r) => {
      const body = r.body as { internetMessageHeaders?: ReadonlyArray<{ name: string; value: string }> } | undefined
      if (r.status >= 200 && r.status < 300 && body?.internetMessageHeaders) {
        unsubscribeBySender.set(
          r.id,
          body.internetMessageHeaders.some((h) => h.name.toLowerCase() === "list-unsubscribe"),
        )
      }
    })
  }, Promise.resolve())

  const decided = messages.map((m) => ({ m, d: classify(m, rules, recordWord, unsubscribeBySender) }))

  // Per-sender report: the approval artefact.
  type Row = {
    address: string
    name: string
    count: number
    reasons: Map<string, number>
    newest: string
    subjects: string[]
  }
  const rows = new Map<string, Row>()
  decided.forEach(({ m, d }) => {
    const a = addressOf(m)
    const row = rows.get(a) ?? {
      address: a,
      name: m.from?.emailAddress.name ?? "",
      count: 0,
      reasons: new Map(),
      newest: "",
      subjects: [] as string[],
    }
    row.count += 1
    row.reasons.set(d.reason, (row.reasons.get(d.reason) ?? 0) + 1)
    const day = m.receivedDateTime?.slice(0, 10) ?? ""
    if (day > row.newest) row.newest = day
    if (row.subjects.length < 3) row.subjects = [...row.subjects, m.subject ?? ""]
    rows.set(a, row)
  })
  const report = [...rows.values()]
    .sort((x, y) => y.count - x.count)
    .map((r) => [
      r.address,
      r.name,
      r.count,
      [...r.reasons.entries()].map(([k, v]) => `${k}=${v}`).join(" "),
      unsubscribeBySender.has(r.address) ? (unsubscribeBySender.get(r.address) ? "yes" : "no") : "",
      r.newest,
      r.subjects.join(" | "),
    ])
  writeFileSync(
    join(outDir, "report.csv"),
    csv([["address", "name", "count", "decisions", "list_unsubscribe", "newest", "sample_subjects"], ...report]),
  )

  const toDelete = decided.filter(({ d }) => d.action === "delete")
  writeFileSync(
    join(outDir, "delete.csv"),
    csv([
      ["received", "address", "subject", "reason", "id"],
      ...toDelete.map(({ m, d }) => [m.receivedDateTime ?? "", addressOf(m), m.subject ?? "", d.reason, m.id]),
    ]),
  )

  const tally = new Map<string, number>()
  decided.forEach(({ d }) => tally.set(d.reason, (tally.get(d.reason) ?? 0) + 1))
  console.log(`\n# Sweep ${apply ? "APPLY" : "rehearsal"} — ${rules.folder} in ${rules.mailbox ?? "own mailbox"}`)
  console.log(`messages: ${messages.length}   delete: ${toDelete.length}   keep: ${messages.length - toDelete.length}`)
  ;[...tally.entries()].sort((a, b) => b[1] - a[1]).forEach(([k, v]) => console.log(`  ${k.padEnd(28)} ${v}`))
  console.log(`\nreport: ${join(outDir, "report.csv")}\ndelete list: ${join(outDir, "delete.csv")}`)

  if (!apply) {
    console.log(`\nNothing moved. Re-run with --apply to move the delete set to ${rules.destination}.`)
    return
  }

  console.error(
    `[sweep] moving ${toDelete.length} messages to ${rules.destination}${paceMs ? ` (pace ${paceMs} ms)` : ""} …`,
  )
  const byId = new Map(toDelete.map(({ m }) => [m.id, m]))
  // moved.csv is appended after every batch, not written at the end: if the run dies
  // part-way, the record of what already moved must survive.
  const movedPath = join(outDir, "moved.csv")
  writeFileSync(movedPath, `${csv([["received", "address", "subject", "id"]])}\n`)
  const result = await runMoveBatches(
    toDelete.map(({ m }) => m.id),
    prefix,
    rules.destination,
    (requests) => client.request<GraphBatchResponse>("POST", "/$batch", { body: { requests } }),
    undefined,
    (done, total, failed, chunkResult) => {
      if (chunkResult.moved.length > 0) {
        const lines = chunkResult.moved.map((id) => {
          const m = byId.get(id)
          return [m?.receivedDateTime ?? "", m ? addressOf(m) : "", m?.subject ?? "", id]
        })
        appendFileSync(movedPath, `${csv(lines)}\n`)
      }
      // Every 200 messages and at the end: enough to see it is alive without a wall of lines.
      if (done % 200 === 0 || done === total) console.error(`[sweep] ${done}/${total} processed, ${failed} failed`)
    },
    paceMs,
  )
  console.log(`moved: ${result.moved.length}   failed: ${result.failed.length}`)
  result.failed.slice(0, 20).forEach((f) => console.log(`  FAILED ${byId.get(f.id)?.subject ?? f.id}: ${f.error}`))
  console.log(`moved list: ${join(outDir, "moved.csv")}`)
}

main().catch((error: unknown) => fail(error instanceof Error ? error.message : String(error)))
