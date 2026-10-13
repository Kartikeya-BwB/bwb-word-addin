import { useEffect, useState } from 'react'
import type { AutoCheck, ChecklistItem } from '../../brand/checklists.ts'
import { AUTHOR_SELF_CHECK, COMPLIANCE_CHECKS, EU_GRANT_CHECKS, WORD_FILE_CHECKS } from '../../brand/checklists.ts'
import { findDocType } from '../../brand/docTypes.ts'
import type { CheckReport } from '../../check/types.ts'
import type { DocInfo } from '../../word/properties.ts'
import { readChecklist, writeChecklist } from '../../word/properties.ts'
import { hasWord } from '../../word/run.ts'
import { Checkbox, Disclosure } from '../components.tsx'
import { useNotify } from '../toast.tsx'

/** The Pre-QA rules behind each automatic checklist item. */
const AUTO_RULES: Record<AutoCheck, RegExp> = {
  spelling: /^writing\/us-spelling/,
  consistency: /^(?:consistency\/hyphenation|structure\/(?:skipped-level|fake-heading|empty-paragraph|heading-case))/,
  acronyms: /^consistency\/(?:acronym|abbreviations)/,
  captions: /^tables\/(?:no-caption|caption-above)/,
  properties: /^send\/(?:no-title|properties)/,
  placeholders: /^(?:send\/(?:placeholder|writing-prompts)|writing\/placeholder|brand\/highlight)/,
  styles: /^(?:brand\/(?:not-template|font|document-font)|tables\/style)/,
  toc: /^structure\/no-contents/,
  'header-footer': /^send\/page-numbers/,
  'tracked-changes': /^send\/(?:tracked-changes|comments)/,
  'alt-text': /^tables\/alt-text/,
  'file-name': /^send\/file-name/,
}

function autoStatus(item: ChecklistItem, report: CheckReport | null): { open: number } | null {
  if (!item.auto || !report) return null
  const rule = AUTO_RULES[item.auto]
  return { open: report.findings.filter((f) => !f.advisory && rule.test(f.rule)).length }
}

function Items({ items, ticks, report, onTick }: {
  items: ChecklistItem[]
  ticks: Record<string, boolean>
  report: CheckReport | null
  onTick: (id: string, value: boolean) => void
}) {
  return (
    <>
      {items.map((item) => {
        const status = autoStatus(item, report)
        const label = (
          <>
            {item.text}
            {status && (
              <span className={`badge ${status.open ? 'minor' : 'ok'}`} style={{ marginLeft: 6 }}>
                {status.open ? `${status.open} open` : 'Check passes'}
              </span>
            )}
          </>
        )
        return <Checkbox key={item.id} label={label} checked={Boolean(ticks[item.id])} onChange={(v) => onTick(item.id, v)} />
      })}
    </>
  )
}

/** The author's self-check before peer review, saved with the document so ticks survive closing it. */
export function Checklist({ report, info }: { report: CheckReport | null; info: DocInfo | null }) {
  const notify = useNotify()
  const [ticks, setTicks] = useState<Record<string, boolean>>({})

  useEffect(() => {
    if (hasWord()) setTicks(readChecklist())
  }, [])

  const docType = findDocType(info?.docType || undefined)
  const typeItems: ChecklistItem[] = (docType?.checks ?? []).map((text, i) => ({ id: `type:${docType?.id ?? ''}:${i}`, text }))
  const groups: Array<{ title: string; items: ChecklistItem[]; open?: boolean }> = [
    { title: 'Author self-check', items: AUTHOR_SELF_CHECK, open: true },
    { title: 'Word file', items: WORD_FILE_CHECKS },
    ...(typeItems.length && docType ? [{ title: `${docType.name} checks`, items: typeItems }] : []),
    ...(info?.funder === 'eu-grant' ? [{ title: 'EU grant deliverable', items: EU_GRANT_CHECKS }] : []),
    { title: 'Compliance, if it could be a financial promotion', items: COMPLIANCE_CHECKS },
  ]
  const all = groups.flatMap((g) => g.items)
  const done = all.filter((i) => ticks[i.id]).length

  const onTick = (id: string, value: boolean): void => {
    const next = { ...ticks, [id]: value }
    setTicks(next)
    if (hasWord()) writeChecklist(next).catch(() => notify('The tick could not be saved with the document.', 'error'))
  }

  return (
    <>
      <p className="muted small" style={{ marginTop: 0 }}>
        {done} of {all.length} ticked. Ticks are saved with the document. Items marked by the check update each time it runs.
      </p>
      {groups.map((g) => (
        <Disclosure key={g.title} title={`${g.title} (${g.items.filter((i) => ticks[i.id]).length}/${g.items.length})`} {...(g.open ? { open: true } : {})}>
          <Items items={g.items} ticks={ticks} report={report} onTick={onTick} />
        </Disclosure>
      ))}
    </>
  )
}
