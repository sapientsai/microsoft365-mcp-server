/** A caller-chosen string, quoted and cut short so it can't add lines to a message. */
export const quoteForMessage = (value: string, max = 60): string => JSON.stringify(value.slice(0, max))
