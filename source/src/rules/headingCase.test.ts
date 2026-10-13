import { describe, expect, it } from 'vitest'
import { checkHeadingCase } from './headingCase.ts'

describe('checkHeadingCase', () => {
  it.each([
    ['Theory Of Change And Impact Measurement', 'Theory of change and impact measurement'],
    ['Annual Report 2025: Key Findings And Next Steps', 'Annual report 2025: Key findings and next steps'],
    ['How UNDP And The SDGs Shape Climate Finance', 'How UNDP and the SDGs shape climate finance'],
    ['Long-Term Finance For Small-Scale Farmers', 'Long-term finance for small-scale farmers'],
    ['1. Introduction To Green Bonds And Climate Finance', '1. Introduction to green bonds and climate finance'],
    ['Why \u201CNet Zero\u201D Matters (And What It Means)', 'Why \u201Cnet zero\u201D matters (and what it means)'],
    ['The Role of the World Bank', 'The role of the world bank'],
    ['Working With BwB And McKinsey On COP28 Targets', 'Working with BwB and McKinsey on COP28 targets'],
    ['Option A Versus Option B', 'Option A versus option B'],
    ['What I Learned About Blended Finance', 'What I learned about blended finance'],
  ])('flags %j and suggests sentence case', (text, suggestion) => {
    expect(checkHeadingCase(text)).toEqual({ titleCase: true, suggestion })
  })

  it.each([
    'Theory of change',
    'Climate finance in East Africa',
    'Annex 2: Terms Of Reference',
    'EXECUTIVE SUMMARY',
    'Working With iPhone Apps Today',
    'UNDP, SDGs And COP28',
    '2025',
    '',
    '   ',
  ])('leaves %j alone', (text) => {
    expect(checkHeadingCase(text)).toEqual({ titleCase: false })
  })

  it('needs at least 75 per cent of the eligible words capitalised', () => {
    // Eligible words: Bonds, Blended, Finance, farmers (3 of 4 capitalised); then Bonds, blended, Finance (2 of 3).
    expect(checkHeadingCase('Green Bonds, Blended Finance and farmers')).toEqual({
      titleCase: true, suggestion: 'Green bonds, blended finance and farmers',
    })
    expect(checkHeadingCase('Green Bonds and blended Finance').titleCase).toBe(false)
  })
})
