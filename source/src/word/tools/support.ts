// Which power tools this Word can run. WordApi sets run on the web, Windows, Mac and iPad; WordApiDesktop sets never
// run on the web, and 1.3 and later are not on iPad either. The pane disables a tool with the reason given here.

import { hasWord, supports } from '../run.ts'

/** True when the host supports the given WordApiDesktop requirement set. Never true on the web. */
export function supportsDesktop(version: string): boolean {
  if (!hasWord()) return false
  try {
    return Office.context.requirements.isSetSupported('WordApiDesktop', version)
  } catch {
    return false
  }
}

/** True in Word on the web. */
export function onTheWeb(): boolean {
  try {
    return typeof Office !== 'undefined' && Office.context.platform === Office.PlatformType.OfficeOnline
  } catch {
    return false
  }
}

/** True in Word for Windows or Mac. */
function desktopPlatform(): boolean {
  try {
    const p = Office.context.platform
    return p === Office.PlatformType.PC || p === Office.PlatformType.Mac
  } catch {
    return false
  }
}

export type Need = 'core' | 'fitToContents' | 'crossReference' | 'tracking' | 'comments' | 'trackedChanges'

const DESKTOP = 'Needs Word for Windows or Mac.'
const newer = (set: string): string => `Needs a newer Word (${set}): update Microsoft 365.`

/** Why a tool cannot run in this Word, or null when it can. In a plain browser nothing is disabled. */
export function unavailable(need: Need): string | null {
  if (!hasWord()) return null
  switch (need) {
    case 'core':
      return supports('1.3') ? null : newer('WordApi 1.3')
    case 'fitToContents':
      return supportsDesktop('1.4') ? null : desktopPlatform() ? newer('WordApiDesktop 1.4') : DESKTOP
    case 'crossReference':
      if (onTheWeb()) return 'Cross-references need Word for Windows or Mac: Word on the web cannot add fields.'
      if (!supports('1.5')) return newer('WordApi 1.5')
      return supportsDesktop('1.1') || desktopPlatform() ? null : DESKTOP
    case 'tracking':
    case 'comments':
      return supports('1.4') ? null : newer('WordApi 1.4')
    case 'trackedChanges':
      return supports('1.6') || supportsDesktop('1.4') ? null : newer('WordApi 1.6')
  }
}
