// Document properties: the core ones Word shows (title, author, company) and BwB custom properties that
// remember the document type, project and deliverable, so checks and file names can use them.

import { runWord } from './run.ts'

export type FunderFrame = 'bwb' | 'eu-grant' | 'un-mdb' | 'uk-public' | 'foundation' | 'private'

export const FUNDER_FRAMES: Array<{ id: FunderFrame; label: string }> = [
  { id: 'bwb', label: 'BwB own (bids, thought leadership, internal)' },
  { id: 'eu-grant', label: 'EU grant' },
  { id: 'un-mdb', label: 'UN or MDB contract' },
  { id: 'uk-public', label: 'UK public sector' },
  { id: 'foundation', label: 'Foundation' },
  { id: 'private', label: 'Private client or BwB venture' },
]

export interface DocInfo {
  title: string
  subject: string
  author: string
  company: string
  docType: string
  project: string
  deliverable: string
  version: string
  funder: FunderFrame | ''
  clientTemplate: boolean
}

const CUSTOM: Record<'docType' | 'project' | 'deliverable' | 'version' | 'funder' | 'clientTemplate', string> = {
  docType: 'BwB document type',
  project: 'BwB project',
  deliverable: 'BwB deliverable',
  version: 'BwB version',
  funder: 'BwB funder',
  clientTemplate: 'BwB client template',
}

function asFunder(value: string): FunderFrame | '' {
  return FUNDER_FRAMES.some((f) => f.id === value) ? (value as FunderFrame) : ''
}

export async function readDocInfo(): Promise<DocInfo> {
  return runWord(async (context) => {
    const props = context.document.properties
    props.load('title,subject,author,company')
    const custom = props.customProperties
    custom.load('items/key,items/value')
    await context.sync()
    const get = (key: string): string => {
      const item = custom.items.find((c) => c.key === key)
      return item ? String(item.value ?? '') : ''
    }
    return {
      title: props.title,
      subject: props.subject,
      author: props.author,
      company: props.company,
      docType: get(CUSTOM.docType),
      project: get(CUSTOM.project),
      deliverable: get(CUSTOM.deliverable),
      version: get(CUSTOM.version),
      funder: asFunder(get(CUSTOM.funder)),
      clientTemplate: get(CUSTOM.clientTemplate) === 'yes',
    }
  })
}

export async function writeDocInfo(info: Partial<DocInfo>): Promise<void> {
  await runWord(async (context) => {
    const props = context.document.properties
    if (info.title !== undefined) props.title = info.title
    if (info.subject !== undefined) props.subject = info.subject
    if (info.author !== undefined) props.author = info.author
    if (info.company !== undefined) props.company = info.company
    const custom = props.customProperties
    const updates: Array<[string, string]> = []
    const put = (key: string, value: string | undefined): void => {
      if (value !== undefined) updates.push([key, value])
    }
    put(CUSTOM.docType, info.docType)
    put(CUSTOM.project, info.project)
    put(CUSTOM.deliverable, info.deliverable)
    put(CUSTOM.version, info.version)
    put(CUSTOM.funder, info.funder)
    if (info.clientTemplate !== undefined) put(CUSTOM.clientTemplate, info.clientTemplate ? 'yes' : '')

    // add() creates or overwrites; an emptied value removes the property instead.
    const removals = updates
      .filter(([, value]) => value === '')
      .map(([key]) => custom.getItemOrNullObject(key))
    removals.forEach((item) => item.load('isNullObject'))
    await context.sync()
    removals.filter((item) => !item.isNullObject).forEach((item) => item.delete())
    updates.filter(([, value]) => value !== '').forEach(([key, value]) => custom.add(key, value))
    await context.sync()
  })
}

/** The file name Word reports for the open document, or null for a new, unsaved one. */
export function documentFileName(): string | null {
  try {
    const url = Office.context.document.url
    if (!url) return null
    const last = decodeURIComponent(url.split(/[\\/]/).pop() ?? '')
    return last.split('?')[0] || null
  } catch {
    return null
  }
}

const CHECKLIST_KEY = 'bwbChecklist'

/** Ticks on the author checklist, stored with the document (Office add-in settings). */
export function readChecklist(): Record<string, boolean> {
  try {
    const value: unknown = Office.context.document.settings.get(CHECKLIST_KEY)
    return value && typeof value === 'object' ? (value as Record<string, boolean>) : {}
  } catch {
    return {}
  }
}

export function writeChecklist(ticks: Record<string, boolean>): Promise<void> {
  return new Promise((resolve, reject) => {
    try {
      Office.context.document.settings.set(CHECKLIST_KEY, ticks)
      Office.context.document.settings.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve()
        else reject(new Error(result.error.message))
      })
    } catch (error) {
      reject(error instanceof Error ? error : new Error(String(error)))
    }
  })
}
