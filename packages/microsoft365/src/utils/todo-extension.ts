// A caller's own data on a To Do task, stored as a Graph open extension. The person can't see or
// edit it, unlike the notes, so a stray edit can't break it.
//
// Checked live on 2026-10-30, in a throwaway list:
// - an extension can be created inside the task's own POST;
// - $expand=extensions($filter=id eq '<bare name>') returns it, and a name with no match returns
//   the task with no extensions key;
// - a PATCH replaces the whole extension, so a merge has to be read-merge-write here;
// - a read carries extensionName, id ("microsoft.graph.openTypeExtension.<name>"), the fields, and
//   per-field annotations such as "rank@odata.type": "#Int64".

/** A field value Graph stores as-is. Nested objects, arrays and null are refused. */
export type ExtensionValue = string | number | boolean

export type ExtensionInput = {
  readonly name: string
  readonly fields: Readonly<Record<string, ExtensionValue>>
}

/** An extension as Graph returns it: its name, its id, the fields, and annotations. */
export type GraphOpenExtension = Readonly<Record<string, unknown>> & { readonly extensionName?: string }

const OPEN_EXTENSION_TYPE = "microsoft.graph.openTypeExtension"

export const EXTENSION_MAX_FIELDS = 20
export const EXTENSION_MAX_BYTES = 2048

// Reverse DNS, e.g. com.civala.assistant. The name also goes inside a quoted $filter, so the
// pattern leaves no room for a quote or a space.
const EXTENSION_NAME = /^[A-Za-z][A-Za-z0-9-]*(?:\.[A-Za-z0-9][A-Za-z0-9-]*)+$/
const EXTENSION_NAME_MAX = 100
// A key Graph would read as an annotation or a path (an @ or a .) can't round-trip as a field.
const FIELD_KEY = /^[A-Za-z][A-Za-z0-9_]*$/
const RESERVED_KEYS = new Set(["extensionname", "id"])

export const extensionNameProblem = (name: string): string | undefined =>
  name.length <= EXTENSION_NAME_MAX && EXTENSION_NAME.test(name)
    ? undefined
    : `Extension name "${name.slice(0, EXTENSION_NAME_MAX)}" must be reverse DNS, such as com.example.app, and at most ${EXTENSION_NAME_MAX} characters.`

const valueProblem = (key: string, value: unknown): string | undefined => {
  if (typeof value === "string" || typeof value === "boolean") return undefined
  if (typeof value === "number") {
    return Number.isFinite(value) ? undefined : `Extension field "${key}" must be a finite number.`
  }
  return `Extension field "${key}" must be a string, number or boolean, not ${value === null ? "null" : typeof value}.`
}

/** Fields that are flat, few, small and safely named; otherwise a message saying what is wrong. */
export const extensionFieldsProblem = (fields: Readonly<Record<string, unknown>>): string | undefined => {
  const entries = Object.entries(fields)
  if (entries.length > EXTENSION_MAX_FIELDS) {
    return `An extension holds at most ${EXTENSION_MAX_FIELDS} fields; this one would hold ${entries.length}.`
  }
  const fieldProblem = entries
    .map(([key, value]) =>
      !FIELD_KEY.test(key) || RESERVED_KEYS.has(key.toLowerCase())
        ? `Extension field name "${key.slice(0, 40)}" must start with a letter and hold only letters, digits and _, and can't be id or extensionName.`
        : valueProblem(key, value),
    )
    .find((problem) => problem !== undefined)
  if (fieldProblem) return fieldProblem
  const bytes = Buffer.byteLength(JSON.stringify(fields), "utf8")
  return bytes > EXTENSION_MAX_BYTES
    ? `An extension holds at most ${EXTENSION_MAX_BYTES} bytes as JSON; this one would hold ${bytes}.`
    : undefined
}

export const extensionProblem = (extension: ExtensionInput): string | undefined =>
  extensionNameProblem(extension.name) ?? extensionFieldsProblem(extension.fields)

/** The body Graph takes to create an extension, inline in a task POST or on its own. */
export const toExtensionBody = (
  name: string,
  fields: Readonly<Record<string, ExtensionValue>>,
): Record<string, unknown> => ({
  "@odata.type": OPEN_EXTENSION_TYPE,
  extensionName: name,
  ...fields,
})

/** The body for a PATCH, which replaces every field, so it carries the whole merged set. */
export const toExtensionPatch = (fields: Readonly<Record<string, ExtensionValue>>): Record<string, unknown> => ({
  "@odata.type": OPEN_EXTENSION_TYPE,
  ...fields,
})

/**
 * Only the caller's fields: Graph's name, id, type and per-field annotations are dropped, and so is
 * any key this server wouldn't accept on input (another writer's), so a key can't break the line.
 */
export const extensionFields = (extension: GraphOpenExtension): Record<string, ExtensionValue> =>
  Object.fromEntries(
    Object.entries(extension).filter(
      (entry): entry is [string, ExtensionValue] =>
        FIELD_KEY.test(entry[0]) &&
        !RESERVED_KEYS.has(entry[0].toLowerCase()) &&
        ["string", "number", "boolean"].includes(typeof entry[1]),
    ),
  )

/** The named extension on a task, matched on extensionName, never on an id prefix. */
export const findExtension = (
  extensions: ReadonlyArray<GraphOpenExtension> | undefined,
  name: string,
): GraphOpenExtension | undefined =>
  (extensions ?? []).find(
    (e) => typeof e.extensionName === "string" && e.extensionName.toLowerCase() === name.toLowerCase(),
  )

/** The $expand entry that returns only the named extension. The name is validated reverse DNS. */
export const extensionExpand = (name: string): string => `extensions($filter=id eq '${name}')`
