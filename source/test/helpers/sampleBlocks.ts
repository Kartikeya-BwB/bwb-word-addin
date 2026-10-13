/** Blocks of every type, for the insertOoxml tests and the round-trip rendering. */
import type { Block } from '../../src/docx/types.ts'
import { picture } from './fixtures.ts'

export function everyBlock(): Block[] {
  return [
    { type: 'heading', level: 1, text: 'Inserted heading {icon:leaf}' },
    { type: 'guidance', text: 'Describe the context for {field:Client name} in two or three sentences.' },
    {
      type: 'paragraph',
      text: 'A paragraph for {field:Client name} with a [link](https://bwb.earth/?a=1&b=2), **bold** text and a note^[With **bold** and *italic* text.].',
    },
    { type: 'lead', text: 'A lead paragraph that introduces the section.' },
    { type: 'accent', text: 'An accent paragraph with one highlighted sentence.' },
    { type: 'bullets', items: ['First action', { text: 'A detail', level: 1 }, 'Second action'], style: 'arrow' },
    { type: 'numbered', items: ['First step', 'Second step'] },
    { type: 'numbered', items: ['A second list restarts at one'] },
    { type: 'pull_quote', text: 'Capital is not the binding constraint.', attribution: 'Bankers without Boundaries' },
    {
      type: 'table',
      header: ['Tier', 'Provider', 'Share'],
      rows: [
        ['First loss', 'Development partners', '10%'],
        ['Senior debt', 'Commercial lenders', '90%'],
        ['Total', '', '100%'],
      ],
      align: [null, null, 'right'],
      total_row: true,
      caption: 'Indicative capital structure.',
      source: 'BwB analysis',
    },
    { type: 'figure', ...picture('media/chart.png'), caption: 'Revenue by pillar.', alt: 'Bar chart of revenue by pillar' },
    {
      type: 'stats',
      hero: 1,
      items: [
        { value: 'USD 250m', label: 'Facility size', icon: 'banknote' },
        { value: '3.2x', label: 'Leverage', icon: 'trending-up' },
        { value: '12', label: 'Communities', icon: 'users' },
      ],
    },
    { type: 'snapshot', items: [{ eyebrow: 'REGION', value: 'Pacific', note: 'Four countries', icon: 'map-pin' }, { eyebrow: 'PERIOD', value: '2024 to 2026' }] },
    { type: 'key_values', rows: [['To', 'Investment Committee'], ['From', 'BwB']] },
    { type: 'callout', title: 'Key message', icon: 'lightbulb', text: 'Coordination, bankability and verification are the constraints.' },
    {
      type: 'case_study',
      title: 'Fisheries cooperative',
      meta: 'Fiji, 2024 to 2026',
      icon: 'fish',
      body: 'A cooperative restored 220 hectares of mangrove.',
      stats: [{ value: '220 ha', label: 'Restored' }],
      source: 'Illustrative',
    },
    {
      type: 'icon_grid',
      columns: 3,
      items: [
        { icon: 'leaf', title: 'Nature finance', text: 'Structures that pay for nature.' },
        { icon: 'landmark', title: 'Capital markets', text: 'Bond structuring.' },
        { icon: 'handshake', title: 'Partnerships', text: 'One set of facts.' },
      ],
    },
    { type: 'caption', kind: 'Figure', text: 'A caption on its own.', source: 'BwB' },
    { type: 'source', text: 'Source: BwB analysis, 2026.' },
    { type: 'toc', levels: 3 },
    { type: 'paragraph', text: 'This paragraph follows the contents, on a new page.' },
    { type: 'disclaimer' },
  ]
}
