import { type Either, Left, Right } from "functype/either"

import type { GraphApiError } from "../types"
import { quoteForMessage } from "./quote"

/**
 * Safe Graph request paths.
 *
 * fetch normalises a URL before sending it: "." and ".." segments are resolved, a backslash reads
 * as "/", tabs and line breaks are dropped, and everything after "#" is cut off. So an ID that
 * goes into a path as written can point the request somewhere else: "A/forward#" turns a
 * createForward into a forward that sends at once, and "../../users/x/messages" leaves the mailbox
 * the server allowed. Two layers stop that:
 *
 *  - `graphPath`, a tagged template, puts each interpolated value in as exactly one path segment.
 *    It refuses an empty ID, ".", "..", or one holding a slash or backslash, and percent-encodes
 *    the rest, so "?", "#", "%" and spaces can't change the URL. Text the code itself built is
 *    marked with `raw()`, so a new path is safe by default.
 *  - `graphPathProblem`, run by core's request on every path, refuses what fetch would rewrite:
 *    a control character, a backslash, a "#", or a dot segment. That makes "the path as written is
 *    the path as sent" true for raw paths too, such as graph_query's.
 */

/** Text the code built itself, put into a path as written. Never wrap a caller's value. */
export type RawPath = { readonly kind: "raw"; readonly text: string }
export const raw = (text: string): RawPath => ({ kind: "raw", text })

/**
 * An ID encoded exactly as encodeURIComponent does, for the paths that always sent it that way
 * (mail folders, inbox rules, extension names). Graph may issue these with a "/" in them
 * (standard base64), which becomes %2F and stays one segment. A "." or ".." piece is refused.
 */
export type EncodedId = { readonly kind: "encoded"; readonly id: string }
export const encodedId = (id: string): EncodedId => ({ kind: "encoded", id })

/**
 * A drive-relative path such as "Documents/Reports", put in as one segment per piece. Empty
 * pieces are dropped, so leading, trailing or doubled slashes don't matter; "." and ".." pieces
 * are refused.
 */
export type DrivePath = { readonly kind: "drive-path"; readonly path: string }
export const drivePath = (path: string): DrivePath => ({ kind: "drive-path", path })

export type GraphPathPart = string | RawPath | EncodedId | DrivePath | Either<GraphApiError, string>

const pathError = (message: string): GraphApiError => ({ type: "api", status: 400, message })

// encodeURIComponent leaves A-Za-z0-9 - _ . ! ~ * ' ( ) alone. These sub-delimiters are also
// legal in a segment and appear in real IDs (site IDs hold commas, chat IDs ":" and "@", base64
// "=" and "+"), so they go back to their literal form and today's wire bytes don't change.
const KEPT_SUB_DELIMS = /%(?:24|26|2B|2C|3A|3B|3D|40)/gi

const encodeSegment = (value: string): Either<string, string> => {
  try {
    return Right(encodeURIComponent(value).replace(KEPT_SUB_DELIMS, (escape) => decodeURIComponent(escape)))
  } catch {
    // encodeURIComponent throws on a lone surrogate: text that isn't well-formed Unicode.
    return Left("isn't well-formed text")
  }
}

const isDotSegment = (value: string): boolean => value === "." || value === ".."

const segmentProblem = (value: string): string | undefined => {
  if (value === "") return "is empty"
  if (isDotSegment(value)) return 'is "." or "..", which would move the request up the path'
  if (/[/\\]/.test(value)) return "holds a slash, which would split it into more than one path segment"
  return undefined
}

const refuse = (value: string, problem: string): Either<GraphApiError, string> =>
  Left(pathError(`The ID ${quoteForMessage(value)} ${problem}. Use an ID exactly as a list or get tool printed it.`))

/** One ID as one encoded path segment, or the reason it can't be one. */
export const idSegment = (id: string): Either<GraphApiError, string> => {
  const problem = segmentProblem(id)
  if (problem !== undefined) return refuse(id, problem)
  return encodeSegment(id).fold<Either<GraphApiError, string>>(
    (reason) => refuse(id, reason),
    (text) => Right(text),
  )
}

const fullyEncodedSegment = (id: string): Either<GraphApiError, string> => {
  if (id === "") return refuse(id, "is empty")
  // encodeURIComponent turns a slash or backslash into %2F or %5C, but a "." or ".." piece would
  // still read as a step up the path to a server that splits on an encoded slash.
  if (id.split(/[/\\]/).some(isDotSegment)) return refuse(id, 'holds a "." or ".." piece')
  try {
    return Right(encodeURIComponent(id))
  } catch {
    return refuse(id, "isn't well-formed text")
  }
}

// ":" ends the path in Graph's root:/...: form, so inside a piece it is encoded. OneDrive and
// SharePoint names can't hold one anyway.
const encodeDrivePiece = (piece: string): Either<string, string> =>
  encodeSegment(piece).map((text) => text.replaceAll(":", "%3A"))

const drivePathSegments = (path: string): Either<GraphApiError, string> => {
  const pieces = path.split("/").filter((piece) => piece !== "")
  if (pieces.length === 0) return Left(pathError("The folder path is empty."))
  const bad = pieces.find((piece) => isDotSegment(piece) || piece.includes("\\"))
  if (bad !== undefined) {
    return Left(pathError(`The folder path can't hold a "." or ".." piece or a backslash: ${quoteForMessage(path)}.`))
  }
  const encoded = pieces.map(encodeDrivePiece)
  const failed = encoded.find((piece) => piece.isLeft())
  if (failed !== undefined) return Left(pathError(`The folder path ${quoteForMessage(path)} isn't well-formed text.`))
  return Right(encoded.map((piece) => piece.value as string).join("/"))
}

const isEither = (part: GraphPathPart): part is Either<GraphApiError, string> =>
  typeof part === "object" && "isLeft" in part

const renderPart = (part: GraphPathPart): Either<GraphApiError, string> => {
  if (typeof part === "string") return idSegment(part)
  if (isEither(part)) return part
  switch (part.kind) {
    case "raw":
      return Right(part.text)
    case "encoded":
      return fullyEncodedSegment(part.id)
    case "drive-path":
      return drivePathSegments(part.path)
  }
}

/**
 * Builds a Graph path. Every interpolated string is an ID and becomes exactly one encoded segment;
 * wrap text the code built in `raw()`. A part that is itself an `Either` (another graphPath) goes
 * in as written, and a `Left` anywhere makes the whole path that `Left`.
 *
 * @example graphPath`${raw(prefix)}/messages/${id}/createForward`
 */
export const graphPath = (
  strings: TemplateStringsArray,
  ...parts: ReadonlyArray<GraphPathPart>
): Either<GraphApiError, string> =>
  parts.reduce<Either<GraphApiError, string>>(
    (built, part, index) =>
      built.flatMap((text) => renderPart(part).map((rendered) => `${text}${rendered}${strings[index + 1] ?? ""}`)),
    Right(strings[0]),
  )

// A segment of only dots, written plainly or percent-encoded, which fetch resolves as "." or "..".
const DOT_SEGMENT = /^(?:\.|%2e){1,2}$/i

/**
 * Why fetch would send this path somewhere other than where it reads, or undefined. Refuses a
 * control character or a "#" anywhere, and a backslash or dot segment before the query string.
 * fetch leaves a backslash in the query alone, and an OData filter may hold one (DOMAIN\\user).
 */
export const graphPathProblem = (path: string): string | undefined => {
  // eslint-disable-next-line no-control-regex -- refusing control characters is the point
  if (/[\u0000-\u001f\u007f]/.test(path)) return "The request path holds a control character, such as a line break."
  if (path.includes("#")) return 'The request path holds a "#", which would cut off the rest of it.'
  const beforeQuery = path.split("?")[0] ?? ""
  if (beforeQuery.includes("\\")) return "The request path holds a backslash, which fetch would read as a slash."
  if (beforeQuery.split("/").some((segment) => DOT_SEGMENT.test(segment))) {
    return 'The request path holds a "." or ".." segment, which would move the request up the path.'
  }
  return undefined
}
