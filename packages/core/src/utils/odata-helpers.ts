import type { ODataParams } from "../types"

// Graph wants the whole $search clause in double quotes, and a double quote or backslash inside it
// escaped with a backslash (https://learn.microsoft.com/graph/search-query-parameter). Unescaped,
// a quoted phrase such as "exact phrase" closes the clause early, and Graph answers
// "An identifier was expected at position 0".
export const escapeSearchClause = (search: string): string => search.replace(/[\\"]/g, (c) => `\\${c}`)

export const buildODataQuery = (params?: ODataParams): string => {
  if (!params) return ""

  const parts: string[] = []

  if (params.$select && params.$select.length > 0) {
    parts.push(`$select=${params.$select.join(",")}`)
  }
  if (params.$filter) {
    parts.push(`$filter=${encodeURIComponent(params.$filter)}`)
  }
  if (params.$expand && params.$expand.length > 0) {
    parts.push(`$expand=${params.$expand.join(",")}`)
  }
  if (params.$orderby) {
    parts.push(`$orderby=${encodeURIComponent(params.$orderby)}`)
  }
  if (params.$top !== undefined) {
    parts.push(`$top=${params.$top}`)
  }
  if (params.$skip !== undefined) {
    parts.push(`$skip=${params.$skip}`)
  }
  if (params.$search) {
    parts.push(`$search="${encodeURIComponent(escapeSearchClause(params.$search))}"`)
  }
  if (params.$count) {
    parts.push(`$count=true`)
  }

  return parts.length > 0 ? `?${parts.join("&")}` : ""
}

// Join a path (which may already carry a query string, e.g. calendarView's
// startDateTime/endDateTime) with the OData query string from buildODataQuery,
// picking the correct separator. Without this, an existing "?" in the path
// turns the appended "?$orderby=..." into part of the previous value, which
// Graph rejects (e.g. "value of parameter EndDateTime is invalid").
export const appendODataQuery = (path: string, queryString: string): string => {
  if (!queryString) return path
  return path.includes("?") ? `${path}&${queryString.slice(1)}` : `${path}${queryString}`
}
