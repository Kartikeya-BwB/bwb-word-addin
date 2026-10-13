// EU grant deliverables carry the EU emblem, a funding acknowledgement and a disclaimer. This is the wording
// most Horizon Europe and LIFE projects use; the grant agreement sets the exact text, so authors check it.

import type { Block } from '../docx/types.ts'

export const EU_ACKNOWLEDGEMENT =
  'Funded by the European Union. Views and opinions expressed are however those of the author(s) only and do not ' +
  'necessarily reflect those of the European Union or the {field:granting authority}. Neither the European Union nor ' +
  'the granting authority can be held responsible for them.'

export function euAcknowledgementBlocks(): Block[] {
  return [
    { type: 'callout', variant: 'charcoal', title: 'Funding acknowledgement', text: EU_ACKNOWLEDGEMENT },
    { type: 'guidance', text: 'Add the EU emblem from the consortium template, name the granting authority and check this wording against the grant agreement.' },
  ]
}
