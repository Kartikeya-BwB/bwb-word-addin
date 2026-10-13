// Links readers outside BwB cannot open: files on someone's computer, shared network folders, intranet addresses and
// the company's own cloud storage. Found in slide hyperlinks and in addresses typed in the text.

export type InternalLink = 'file' | 'network' | 'intranet' | 'storage'

/** Host names of cloud storage that only signed-in staff can usually open. */
const STORAGE_HOSTS = [/\.sharepoint\.com$/i, /^onedrive\.live\.com$/i, /^1drv\.ms$/i]
const INTRANET_SUFFIXES = /\.(?:local|internal|intranet|corp|lan|localdomain)$/i

function privateAddress(host: string): boolean {
  const m = /^(\d{1,3})\.(\d{1,3})\.\d{1,3}\.\d{1,3}$/.exec(host)
  if (!m) return false
  const [a, b] = [Number(m[1]), Number(m[2])]
  return a === 10 || a === 127 || (a === 192 && b === 168) || (a === 172 && b >= 16 && b <= 31)
}

/** What kind of internal place an address points to, or null for a public one. */
export function internalLink(address: string): InternalLink | null {
  const a = address.trim()
  if (/^file:/i.test(a) || /^[a-z]:[\\/]/i.test(a)) return 'file'
  if (/^\\\\[^\\]/.test(a)) return 'network'
  let host: string
  try {
    host = new URL(a).hostname
  } catch {
    return null
  }
  if (!host) return null
  if (STORAGE_HOSTS.some((re) => re.test(host))) return 'storage'
  if (host === 'localhost' || !host.includes('.') || INTRANET_SUFFIXES.test(host) || privateAddress(host)) return 'intranet'
  return null
}

const ADDRESS = /\b(?:https?|file):\/\/[^\s<>"\u201C\u201D\u2018\u2019]+|\\\\[^\s\\]+\\[^\s]*|\b[a-z]:\\[^\s]*/gi

/** Web addresses and file paths typed in a text, with trailing punctuation trimmed. */
export function addressesIn(text: string): string[] {
  return [...text.matchAll(ADDRESS)].map((m) => m[0].replace(/[.,;:)\]]+$/, '')).filter((a) => a.length > 3)
}

export const INTERNAL_LINK_DETAIL: Record<InternalLink, string> = {
  file: 'It points to a file on someone’s computer, which nobody else can open. Link to a public page, or send the file.',
  network: 'It points to a shared network folder that only BwB staff can reach. Link to a public page, or send the file.',
  intranet: 'It points to an internal web address that only works inside BwB. Link to a public page instead.',
  storage: 'It points to BwB’s cloud storage, which readers outside BwB may not be able to open. Check it is shared with them, or send the file.',
}
