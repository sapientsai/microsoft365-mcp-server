import { AsyncLocalStorage } from "node:async_hooks"

const tokenContext = new AsyncLocalStorage<string | undefined>()

export const withToken = <T>(token: string | undefined, fn: () => T): T => tokenContext.run(token, fn)

export const getContextToken = (): string | undefined => tokenContext.getStore()
