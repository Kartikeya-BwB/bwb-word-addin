// Thin wrappers around Word.run so every feature fails the same, readable way.

export class NotInWordError extends Error {
  constructor() {
    super('Open this in Word to use it. The task pane is running in a plain browser.')
    this.name = 'NotInWordError'
  }
}

export class UserFacingError extends Error {
  override name = 'UserFacingError'
}

export function hasWord(): boolean {
  return typeof Word !== 'undefined' && typeof Office !== 'undefined'
}

/** True when the host supports the given WordApi requirement set (1.1 to 1.9). */
export function supports(version: string): boolean {
  if (!hasWord()) return false
  try {
    return Office.context.requirements.isSetSupported('WordApi', version)
  } catch {
    return false
  }
}

export async function runWord<T>(fn: (context: Word.RequestContext) => Promise<T>): Promise<T> {
  if (!hasWord()) throw new NotInWordError()
  return Word.run(fn)
}

interface OfficeErrorLike {
  code?: unknown
  message?: unknown
  debugInfo?: { message?: unknown; errorLocation?: unknown }
}

/** A short, plain message for a toast. Office errors carry a code that is more useful than the message. */
export function errorMessage(error: unknown): string {
  if (error instanceof NotInWordError || error instanceof UserFacingError) return error.message
  const e = (error ?? {}) as OfficeErrorLike
  const code = typeof e.code === 'string' ? e.code : ''
  const message = typeof e.message === 'string' ? e.message : String(error)
  switch (code) {
    case 'ItemNotFound':
      return 'Word could not find that item. The document may have changed: run it again.'
    case 'InvalidArgument':
      return `Word rejected the change (${message}).`
    case 'AccessDenied':
      return 'Word did not allow the change. The document may be read-only or protected.'
    case 'ApiNotFound':
      return 'This version of Word does not support that feature yet. Update Microsoft 365 and try again.'
    default:
      return code ? `${message} (${code})` : message
  }
}
