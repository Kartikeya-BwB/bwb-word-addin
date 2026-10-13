import { describe, expect, it } from 'vitest'
import { changeCase, isShouting } from '../../../src/word/tools/textCase.ts'

describe('changeCase: sentence case', () => {
  it.each([
    ['How UNDP And The SDGs Shape Climate Finance', 'How UNDP and the SDGs shape climate finance'],
    ['Working With BwB And McKinsey On COP28 Targets', 'Working with BwB and McKinsey on COP28 targets'],
    ['THE ROLE OF THE GCF. WHAT COMES NEXT?', 'The role of the GCF. What comes next?'],
    ['what i learned about blended finance', 'What I learned about blended finance'],
    // Names of places cannot be told from other words: the message asks the author to check them.
    ['iPhone Users In Kenya', 'iPhone users in kenya'],
    ['UN-Backed Funds', 'UN-backed funds'],
    ['Long-Term Finance For Small-Scale Farmers', 'Long-term finance for small-scale farmers'],
    ['See e.g. The Annex. Then Read On', 'See e.g. the annex. Then read on'],
    ['\u201cNet Zero\u201d Matters: A Review', '\u201cNet zero\u201d matters: a review'],
    ['EXECUTIVE SUMMARY', 'Executive summary'],
  ])('%j', (text, expected) => {
    expect(changeCase(text, 'sentence')).toBe(expected)
  })

  it('keeps acronyms the document uses when the text is all in capitals', () => {
    expect(changeCase('UNICEF AND THE WATER SECTOR', 'sentence')).toBe('Unicef and the water sector')
    expect(changeCase('UNICEF AND THE WATER SECTOR', 'sentence', { acronyms: new Set(['UNICEF']) })).toBe('UNICEF and the water sector')
  })

  it('writes the company name its own way in text in capitals', () => {
    expect(changeCase('ABOUT BWB', 'sentence')).toBe('About BwB')
  })

  it('keeps the spaces and line breaks between words', () => {
    expect(changeCase('ONE  TWO\tTHREE\u000bFOUR', 'sentence')).toBe('One  two\tthree\u000bfour')
  })
})

describe('changeCase: title case', () => {
  it.each([
    ['the role of the world bank in africa', 'The Role of the World Bank in Africa'],
    ['state-of-the-art finance for UNDP', 'State-of-the-Art Finance for UNDP'],
    ['annex 2: terms of reference', 'Annex 2: Terms of Reference'],
    ['what we are working on', 'What We Are Working On'],
    ['BLENDED FINANCE AND THE GCF', 'Blended Finance and the GCF'],
    ['working with iPhone users', 'Working with iPhone Users'],
  ])('%j', (text, expected) => {
    expect(changeCase(text, 'title')).toBe(expected)
  })
})

describe('changeCase: lower and upper case', () => {
  it('lower case keeps acronyms and names with capitals inside', () => {
    expect(changeCase('The UNDP Report On BwB And McKinsey', 'lower')).toBe('the UNDP report on BwB and McKinsey')
  })

  it('lower case of text in capitals keeps only the acronyms it can tell', () => {
    expect(changeCase('THE MDBS AND THE GCF REPORT', 'lower')).toBe('the MDBS and the GCF report')
  })

  it('upper case changes everything', () => {
    expect(changeCase('BwB and the UNDP', 'upper')).toBe('BWB AND THE UNDP')
  })
})

describe('isShouting', () => {
  it('spots text written in capitals', () => {
    expect(isShouting('EXECUTIVE SUMMARY')).toBe(true)
    expect(isShouting('INTRODUCTION')).toBe(true)
    expect(isShouting('UNDP')).toBe(false)
    expect(isShouting('UNDP and GCF')).toBe(false)
    expect(isShouting('2025')).toBe(false)
  })
})
