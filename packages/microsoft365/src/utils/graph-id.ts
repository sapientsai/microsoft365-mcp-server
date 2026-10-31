// Graph's ids for messages, events and calendars are base64 in its URL-safe form, about 150
// characters (checked live, 2026-10-31). An id goes into a request path, so anything else is
// refused before it reaches Graph: a "/" or "." could resolve the path somewhere else, such as
// /users/{someone}/events, past MS365_ALLOWED_MAILBOXES.
export const isGraphId = (id: string): boolean => /^[A-Za-z0-9_=-]{1,1000}$/.test(id)
