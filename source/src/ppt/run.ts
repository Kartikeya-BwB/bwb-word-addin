// Thin wrappers around PowerPoint.run so every PowerPoint feature fails the same, readable way: the PowerPoint twin
// of src/word/run.ts. Errors become UserFacingError, which the pane's toasts and the ribbon's notice show as they are.

import { inHost } from '../host.ts'
import { UserFacingError } from '../word/run.ts'

export class NotInPowerPointError extends UserFacingError {
  constructor() {
    super('Open this in PowerPoint to use it. The task pane is running in a plain browser.')
  }
}

export function hasPowerPoint(): boolean {
  return typeof PowerPoint !== 'undefined' && inHost('powerpoint')
}

/** True when PowerPoint supports the PowerPointApi requirement set (1.1 to 1.10). */
export function supportsPpt(version: string): boolean {
  if (!hasPowerPoint()) return false
  try {
    return Office.context.requirements.isSetSupported('PowerPointApi', version)
  } catch {
    return false
  }
}

/** What to tell people whose PowerPoint is too old for a feature. */
export function updateMessage(feature: string): string {
  return `${feature} needs a newer PowerPoint. Update Microsoft 365, or use PowerPoint on the web.`
}

/** Why a pane control is unavailable, or null when it can run. In a plain browser everything is available to try. */
export function unavailableReason(version: string, feature: string): string | null {
  if (!hasPowerPoint()) return null
  return supportsPpt(version) ? null : updateMessage(feature)
}

/** Throws a readable error when PowerPoint lacks the requirement set a feature needs. */
export function requirePpt(version: string, feature: string): void {
  if (!hasPowerPoint()) throw new NotInPowerPointError()
  if (!supportsPpt(version)) throw new UserFacingError(updateMessage(feature))
}

interface OfficeErrorLike {
  code?: unknown
  message?: unknown
}

/** A short, plain message for an Office error, in PowerPoint's terms. */
export function pptErrorMessage(error: unknown): string {
  const e = (error ?? {}) as OfficeErrorLike
  const code = typeof e.code === 'string' ? e.code : ''
  const message = typeof e.message === 'string' ? e.message : String(error)
  switch (code) {
    case 'ItemNotFound':
    case 'SlideNotFound':
      return 'PowerPoint could not find that shape or slide. The slide may have changed: try again.'
    case 'InvalidArgument':
      return `PowerPoint rejected the change (${message}).`
    case 'InvalidSelection':
      return 'That does not work with what is selected. Select shapes on a slide and try again.'
    case 'AccessDenied':
      return 'PowerPoint did not allow the change. The presentation may be read-only.'
    case 'ApiNotFound':
      return 'This version of PowerPoint does not support that feature yet. Update Microsoft 365 and try again.'
    default:
      return code ? `${message} (${code})` : message
  }
}

/** Runs a batch against PowerPoint, turning Office errors into readable ones. Programming errors pass through. */
export async function runPpt<T>(fn: (context: PowerPoint.RequestContext) => Promise<T>): Promise<T> {
  if (!hasPowerPoint()) throw new NotInPowerPointError()
  try {
    return await PowerPoint.run(fn)
  } catch (error) {
    if (error instanceof UserFacingError) throw error
    const code = (error as OfficeErrorLike | null)?.code
    if (typeof code === 'string' && code !== 'PropertyNotLoaded') throw new UserFacingError(pptErrorMessage(error))
    throw error
  }
}
