import type { ODataParams } from "../types"

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
    parts.push(`$search="${encodeURIComponent(params.$search)}"`)
  }
  if (params.$count) {
    parts.push(`$count=true`)
  }

  return parts.length > 0 ? `?${parts.join("&")}` : ""
}
