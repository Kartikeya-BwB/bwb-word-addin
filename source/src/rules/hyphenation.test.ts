import { describe, expect, it } from 'vitest'
import { findHyphenationVariants } from './hyphenation.ts'

describe('findHyphenationVariants', () => {
  it('finds hyphenated and closed forms of the same compound', () => {
    expect(findHyphenationVariants(['Send an e-mail today.', 'Check your email and your E-mail settings.'])).toEqual([
      { forms: [{ form: 'e-mail', count: 2 }, { form: 'email', count: 1 }], suggestion: 'e-mail' },
    ])
    expect(findHyphenationVariants(['Co-operation matters.', 'International cooperation and cooperation agreements.'])).toEqual([
      { forms: [{ form: 'cooperation', count: 2 }, { form: 'co-operation', count: 1 }], suggestion: 'cooperation' },
    ])
  })

  it('finds open and closed forms', () => {
    expect(findHyphenationVariants(['The web site is new.', 'Visit our website.'])).toEqual([
      { forms: [{ form: 'website', count: 1 }, { form: 'web site', count: 1 }], suggestion: 'website' },
    ])
  })

  it('counts all three forms case-insensitively', () => {
    const got = findHyphenationVariants([
      'Decision-making is slow.', 'Good decision making helps.', 'Decision making takes time.', 'decisionmaking',
    ])
    expect(got).toEqual([{
      forms: [{ form: 'decision making', count: 2 }, { form: 'decision-making', count: 1 }, { form: 'decisionmaking', count: 1 }],
      suggestion: 'decision making',
    }])
  })

  it('breaks ties with the hyphenated form for adjective compounds, otherwise the closed form', () => {
    expect(findHyphenationVariants(['A long-term plan.', 'In the long term, we grow.'])[0]?.suggestion).toBe('long-term')
    expect(findHyphenationVariants(['Apply on-line.', 'Apply online.'])[0]?.suggestion).toBe('online')
    expect(findHyphenationVariants(['A small-scale farm.', 'A smallscale farm.'])[0]?.suggestion).toBe('small-scale')
  })

  it('ignores numbers, ranges, chains and open forms without another spelling', () => {
    expect(findHyphenationVariants([
      'From 2022-2025, 5-10 per cent of COVID-19 cases.', '2022 2025 and 5 10.', 'A state-of-the-art plant, state of the art.',
      'Decision making alone is fine.', 'Section A-B and a b.',
    ])).toEqual([])
  })

  it('does not pair function words or phrasal verbs with closed forms', () => {
    expect(findHyphenationVariants(['It may be late.', 'Maybe not.', 'Log in to the portal.', 'It goes into force.'])).toEqual([])
    expect(findHyphenationVariants(['Set up the fund.', 'The setup took time.', 'We follow up.', 'A follow-up call.'])).toEqual([])
    // Hyphenated against closed is still compared for such pairs.
    expect(findHyphenationVariants(['A follow-up call.', 'A followup call.'])[0]?.forms.map((f) => f.form).sort())
      .toEqual(['follow-up', 'followup'])
  })

  it('reads non-breaking hyphens and spaces and keeps the spelling as written', () => {
    const got = findHyphenationVariants(['An e\u2011mail.', 'An email.', 'In the long\u00A0term.', 'A long-term view.'])
    expect(got).toEqual([
      { forms: [{ form: 'email', count: 1 }, { form: 'e\u2011mail', count: 1 }], suggestion: 'email' },
      { forms: [{ form: 'long-term', count: 1 }, { form: 'long\u00A0term', count: 1 }], suggestion: 'long-term' },
    ])
  })

  it('lists variants in order of first appearance', () => {
    const got = findHyphenationVariants(['online and on-line', 'long-term and long term', 'e-mail and email'])
    expect(got.map((v) => v.suggestion)).toEqual(['online', 'long-term', 'email'])
  })
})
