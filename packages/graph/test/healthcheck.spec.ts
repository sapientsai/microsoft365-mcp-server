// The image's HEALTHCHECK probes a path on this server. When the server stopped serving it, every
// container went unhealthy about 90 s after starting and the orchestrator killed it, while the old
// task kept serving. So the path is read from the Dockerfile itself, and the server must answer it
// with no API key, the way the probe calls it.
import { readFileSync } from "node:fs"
import { dirname, join } from "node:path"
import { fileURLToPath } from "node:url"

import { type AuthStrategy, mintUploadTicket } from "@sapientsai/ms-graph-core"
import { Right } from "functype/either"
import { describe, expect, it } from "vitest"

import type { ServerRuntimeConfig } from "../src/config"
import { buildServer } from "../src/index"

const here = dirname(fileURLToPath(import.meta.url))
const dockerfile = readFileSync(join(here, "..", "Dockerfile"), "utf-8")

const probedPath = (): string => {
  const match = /HEALTHCHECK[\s\S]*?https?:\/\/127\.0\.0\.1:\$\{PORT\}(\/[^\s|"']*)/.exec(dockerfile)
  if (!match?.[1]) throw new Error("No HEALTHCHECK URL found in packages/graph/Dockerfile")
  return match[1]
}

const auth: AuthStrategy = { getAccessToken: () => Promise.resolve(Right("TOKEN")) }

const config = (apiKey?: string): ServerRuntimeConfig => ({
  auth: { tenantId: "tenant", clientId: "client", clientSecret: "made-up", appScopes: [] },
  ...(apiKey ? { apiKey } : {}),
  transport: "httpStream",
  port: 3000,
  host: "127.0.0.1",
  publicBaseUrl: "http://127.0.0.1:3000",
  enableDownloadFile: false,
  sharePointSearch: { region: "NAM" },
})

// A compose healthcheck replaces the image's, so it must probe the same path, on IPv4: in an Alpine
// container "localhost" can resolve to ::1 first, and the server binds IPv4 only.
const composeProbe = (): string | undefined =>
  /healthcheck:[\s\S]*?"(https?:\/\/[^"]+)"/.exec(readFileSync(join(here, "..", "docker-compose.yml"), "utf-8"))?.[1]

describe("the compose healthcheck", () => {
  it("probes the same path as the image, on 127.0.0.1", () => {
    expect(composeProbe()).toBe(`http://127.0.0.1:8080${probedPath()}`)
  })
})

// The ticket get_upload_config hands out is printed into the transcript, so it must not open /mcp.
describe("an upload ticket on the MCP endpoint", () => {
  it("is refused, while the API key itself is accepted", async () => {
    const server = buildServer(config("made-up-key"), auth)
    const ticket = mintUploadTicket("made-up-key", { relay: "drive", path: "/me/drive/root:/a.txt:/content" })
    const post = (bearer: string) =>
      server.fetch(
        new Request("http://127.0.0.1:3000/mcp", {
          method: "POST",
          headers: {
            authorization: `Bearer ${bearer}`,
            accept: "application/json, text/event-stream",
            "content-type": "application/json",
          },
          body: JSON.stringify({ jsonrpc: "2.0", id: 1, method: "tools/list", params: {} }),
        }),
      )
    expect((await post(ticket)).status).toBe(401)
    expect((await post("made-up-key")).status).not.toBe(401)
  })
})

describe("the Docker HEALTHCHECK", () => {
  it("probes /ping", () => {
    expect(probedPath()).toBe("/ping")
  })

  it.each([
    ["with an API key configured", "made-up-key"],
    ["with no API key", undefined],
  ])("is answered with 200 and no credentials, %s", async (_, apiKey) => {
    const server = buildServer(config(apiKey), auth)
    const response = await server.fetch(new Request(`http://127.0.0.1:3000${probedPath()}`))
    expect(response.status).toBe(200)
    expect(await response.text()).toBe("pong")
  })
})
