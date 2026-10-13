// The building-block library: each block arrives with grey placeholders to type over, which the Check tab lists
// until they are filled. Used by the Insert tab and by the ribbon's Blocks menu. Section dividers are whole Word
// sections, which cannot be inserted into an open document, so they are an option when starting a report instead.

import type { Block } from '../docx/types.ts'

export interface LibraryBlock {
  id: string
  label: string
  hint: string
  blocks: () => Block[]
}

const three = [1, 2, 3]

export const LIBRARY: LibraryBlock[] = [
  {
    id: 'callout', label: 'Callout', hint: 'A takeaway, tip or lesson, with a green bar',
    blocks: () => [{ type: 'callout', title: '{field:Title}', text: '{field:Key message}' }],
  },
  {
    id: 'note', label: 'Note', hint: 'A caveat or an assumption, with a charcoal bar',
    blocks: () => [{ type: 'callout', variant: 'charcoal', title: '{field:Note}', text: '{field:Caveat or assumption}' }],
  },
  {
    id: 'stats', label: 'Key figures', hint: 'Two to four headline numbers on navy',
    blocks: () => [{ type: 'stats', items: three.map((n) => ({ value: `{field:Figure ${n}}`, label: '{field:Label}' })), hero: 0 }],
  },
  {
    id: 'snapshot', label: 'Snapshot', hint: 'A navy grid of key facts to open a briefing',
    blocks: () => [{ type: 'snapshot', items: three.map(() => ({ eyebrow: '{field:Eyebrow}', value: '{field:Value}', note: '{field:Note}' })) }],
  },
  {
    id: 'quote', label: 'Pull quote', hint: 'One quotation that carries a point',
    blocks: () => [{ type: 'pull_quote', text: '{field:Quotation}', attribution: '{field:Name, role, organisation}' }],
  },
  {
    id: 'case', label: 'Case study', hint: 'A titled box with a row of figures and a source',
    blocks: () => [{
      type: 'case_study', title: '{field:Case study title}', meta: '{field:Country, year}', body: '{field:What happened and why it matters}',
      stats: [{ value: '{field:Figure}', label: '{field:Label}' }], source: '{field:Source}',
    }],
  },
  {
    id: 'grid', label: 'Icon grid', hint: 'Services, workstreams or principles, each with an icon',
    blocks: () => [{
      type: 'icon_grid', columns: 3,
      items: ['target', 'handshake', 'trending-up'].map((icon) => ({ icon, title: '{field:Title}', text: '{field:One sentence}' })),
    }],
  },
  {
    id: 'details', label: 'Details table', hint: 'Label and value pairs, such as a project summary',
    blocks: () => [{ type: 'key_values', rows: [['Client', '{field:Name}'], ['Period', '{field:Dates}'], ['Budget', '{field:Amount}']] }],
  },
]

export function findLibraryBlock(id: string): LibraryBlock | undefined {
  return LIBRARY.find((b) => b.id === id)
}

/** A four-column BwB data table with a caption placeholder, as the ribbon's Table button inserts it. */
export function defaultTableBlocks(): Block[] {
  return [{
    type: 'table', style: 'standard',
    header: [1, 2, 3, 4].map((n) => `{field:Column ${n}}`),
    rows: Array.from({ length: 4 }, () => [null, null, null, null]),
    caption: '{field:Table title}',
  }]
}
