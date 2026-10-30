// The characters Graph's To Do delta tokens were seen to use (live, 2026-10-30). A token is the one
// value a caller hands back, so anything else is refused rather than put into a URL.
export const isDeltaToken = (token: string): boolean => token.length <= 4096 && /^[A-Za-z0-9._~-]+$/.test(token)
