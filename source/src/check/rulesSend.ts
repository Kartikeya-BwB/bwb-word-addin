// Before you send: tracked changes, comments, prompts and placeholders, stickers, properties, footer, file name.

import { fileNameIssues } from './fileName.ts'
import type { DocSnapshot, Finding } from './types.ts'
import { plural } from './util.ts'

const BWB = 'Bankers without Boundaries'

export function sendFindings(snapshot: DocSnapshot): Finding[] {
  const out: Finding[] = []
  if (snapshot.trackedChanges) {
    out.push({
      id: 'send:tracked', rule: 'send/tracked-changes', group: 'send', severity: 'major',
      title: `${plural(snapshot.trackedChanges, 'tracked change')} still in the document`,
      detail: 'Accept or reject them (Review, Accept) before the document goes to QA or the client.',
    })
  }
  if (snapshot.comments) {
    out.push({
      id: 'send:comments', rule: 'send/comments', group: 'send', severity: 'major',
      title: `${plural(snapshot.comments, 'comment')} in the document`,
      detail: 'Resolve and delete comments before sending (Review, Delete, Delete All Comments).',
    })
  }

  const prompts = snapshot.controls.filter((c) => c.tag === 'BwBGuidance')
  if (prompts.length) {
    out.push({
      id: 'send:prompts', rule: 'send/writing-prompts', group: 'send', severity: 'major',
      title: `${plural(prompts.length, 'writing prompt')} left in`,
      detail: 'Prompts from the BwB template guide drafting. Replace or delete each one before sending.',
      location: { kind: 'control', control: prompts[0]?.index ?? 0, tag: 'BwBGuidance' },
      fix: { kind: 'remove-controls', tag: 'BwBGuidance' }, fixLabel: 'Delete all prompts',
    })
  }
  for (const c of snapshot.controls.filter((x) => x.tag === 'BwBPlaceholder' && /^\[.*\]$/.test(x.text.trim()))) {
    out.push({
      id: `send:placeholder:${c.index}`, rule: 'send/placeholder', group: 'send', severity: 'major',
      title: `Not filled in: ${c.text.trim()}`, location: { kind: 'control', control: c.index, tag: 'BwBPlaceholder' },
    })
  }
  const sticker = snapshot.controls.find((c) => c.tag === 'BwBSticker')
  if (sticker) {
    out.push({
      id: 'send:sticker', rule: 'send/sticker', group: 'send', severity: 'minor', advisory: true,
      title: `“${sticker.text.trim()}” sticker in the header`, detail: 'Remove it before sending a final version.',
      fix: { kind: 'remove-sticker' }, fixLabel: 'Remove the sticker',
    })
  }

  const props = snapshot.info
  if (!props.title.trim()) {
    out.push({
      id: 'send:title', rule: 'send/no-title', group: 'send', severity: 'minor',
      title: 'The document title property is empty', detail: 'Word shows it in search, PDFs and SharePoint. Set it on the Start tab.',
    })
  }
  if (props.company !== BWB || props.author !== BWB) {
    out.push({
      id: 'send:properties', rule: 'send/properties', group: 'send', severity: 'copyedit',
      title: 'Author and company properties not cleaned',
      detail: `They read “${props.author || 'empty'}” and “${props.company || 'empty'}”. Client documents carry ${BWB}.`,
      fix: { kind: 'properties', values: { author: BWB, company: BWB } }, fixLabel: `Set both to ${BWB}`,
    })
  }

  const footers = snapshot.headersFooters.filter((h) => h.kind === 'footer')
  if (footers.length && !footers.some((f) => /page\s+\d+\s+of\s+\d+/i.test(f.text))) {
    out.push({
      id: 'send:page-numbers', rule: 'send/page-numbers', group: 'send', severity: 'minor',
      title: 'No “Page X of Y” in the footer', detail: 'BwB documents number pages bottom right. Documents started from the BwB template have it.',
    })
  }

  if (snapshot.fileName) {
    for (const issue of fileNameIssues(snapshot.fileName)) {
      out.push({
        id: `send:file:${issue.code}`, rule: `send/file-name-${issue.code}`, group: 'send', severity: 'minor',
        advisory: issue.code === 'final', title: issue.message, excerpt: snapshot.fileName,
      })
    }
  }
  return out
}
