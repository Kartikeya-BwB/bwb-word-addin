// File-name hygiene from the BwB document-type research (section 8): internal markers keep leaking into
// finals, so they are checked, and a clean name is suggested from the document's properties.

export interface FileNameIssue {
  code: string
  message: string
}

const RULES: Array<{ code: string; test: RegExp; message: string }> = [
  { code: 'copy', test: /(^|[\s_\-(])(copy|kopie)(\s*\(\d+\))?($|[\s_\-).])/i, message: 'Remove “Copy” from the file name' },
  { code: 'copy-number', test: /\(\d+\)|\[\d+\]/, message: 'Remove the copy number such as “(1)”' },
  { code: 'old', test: /(^|[\s_\-])OLD($|[\s_\-.])/, message: 'Remove “OLD” from the file name' },
  { code: 'tracked', test: /_TC($|[\s_\-.])/, message: 'Remove the “_TC” tracked-changes marker' },
  { code: 'comments', test: /(^|[\s_\-])comments?($|[\s_\-.])/i, message: 'Remove “Comments” from the file name' },
  { code: 'initials', test: /(^|[\s_])[A-Z]{2}\s+edits($|[\s_.])|_[A-Z]{2}(\s[A-Z]{2})+$/, message: 'Remove reviewer initials' },
  { code: 'internal', test: /internal[\s_\-]*only/i, message: 'Remove “internal only” before the file leaves BwB' },
  { code: 'final', test: /(^|[\s_\-])(final|vF)($|[\s_\-.\d])/i, message: 'A file’s state belongs in the hub, so names never need “FINAL” or “vF”' },
]

export function fileNameIssues(fileName: string): FileNameIssue[] {
  const stem = fileName.replace(/\.(docx?|dotx|pdf)$/i, '')
  return RULES.filter((r) => r.test.test(stem)).map(({ code, message }) => ({ code, message }))
}

function clean(part: string): string {
  return part.replace(/[\\/:*?"<>|]+/g, ' ').replace(/\s+/g, ' ').trim()
}

function shortTitle(title: string, words = 6): string {
  return clean(title).split(' ').slice(0, words).join(' ')
}

export interface FileNameParts {
  project?: string
  deliverable?: string
  title?: string
  version?: string
}

/** YYYYMMDD_<Project>_<Deliverable no.>_<Short title>_v<NN>.docx, leaving out parts that are not known. */
export function suggestFileName(parts: FileNameParts, date: Date = new Date()): string {
  const y = date.getFullYear()
  const m = String(date.getMonth() + 1).padStart(2, '0')
  const d = String(date.getDate()).padStart(2, '0')
  const version = (parts.version ?? '').replace(/^v/i, '').trim()
  const segments = [
    `${y}${m}${d}`,
    clean(parts.project ?? ''),
    clean(parts.deliverable ?? ''),
    shortTitle(parts.title ?? '') || 'Untitled',
    `v${(version || '1').padStart(2, '0')}`,
  ].filter(Boolean)
  return `${segments.join('_')}.docx`
}
