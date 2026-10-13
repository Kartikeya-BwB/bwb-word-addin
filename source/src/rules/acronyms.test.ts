import { describe, expect, it } from 'vitest'
import { analyseAcronyms, type AcronymEntry, type AcronymInputParagraph } from './acronyms.ts'

const paras = (...texts: string[]): AcronymInputParagraph[] => texts.map((text) => ({ text }))
const byName = (entries: AcronymEntry[]): Record<string, AcronymEntry> =>
  Object.fromEntries(entries.map((e) => [e.acronym, e]))
const summary = (entries: AcronymEntry[]): Array<[string, string, string | undefined]> =>
  entries.map((e) => [e.acronym, e.status, e.definition])

describe('analyseAcronyms', () => {
  it('reads a realistic BwB passage', () => {
    const doc = paras(
      'The United Nations Development Programme (UNDP) supports SDG bonds in Kenya.',
      'UNDP and the multilateral development banks (MDBs) agreed a common approach.',
      'Each MDB reports to the Net Zero Cities (NZC) mission, as set out in the Grant Agreement (GA).',
      'Deliverable D7.4 and task T2.1 are due in month 18; the GA sets the budget at EUR 1.2m and USD 5m.',
      'See https://www.undp.org/SDG-NZC or write to info@UNDP.org.',
    )
    const got = analyseAcronyms(doc)
    expect(summary(got)).toEqual([
      ['GA', 'defined', 'Grant Agreement'],
      ['MDB', 'defined', 'multilateral development banks'],
      ['NZC', 'defined', 'Net Zero Cities'],
      ['SDG', 'undefined', undefined],
      ['UNDP', 'defined', 'United Nations Development Programme'],
    ])
    const e = byName(got)
    expect(e.UNDP).toMatchObject({ uses: 2, firstUse: { paragraph: 0, start: 42, end: 46 }, definedAt: { paragraph: 0, start: 4, end: 47 } })
    expect(e.MDB).toMatchObject({ uses: 2, firstUse: { paragraph: 1, start: 45, end: 49 } })
    expect(e.SDG).toMatchObject({ uses: 1, firstUse: { paragraph: 0, start: 57, end: 60 } })
    expect(e.GA?.uses).toBe(2)
  })

  it('reports an acronym used before it is spelled out', () => {
    const got = analyseAcronyms(paras('The SDGs guide our work.', 'The Sustainable Development Goals (SDGs) are global.'))
    expect(got).toEqual([{
      acronym: 'SDG', status: 'defined-late', definition: 'Sustainable Development Goals', uses: 2,
      firstUse: { paragraph: 0, start: 4, end: 8 }, definedAt: { paragraph: 1, start: 4, end: 40 },
    }])
  })

  it('accepts the acronym-first form and prefers the earliest definition', () => {
    const got = analyseAcronyms(paras(
      'UNDP (United Nations Development Programme) leads.',
      'Later: United Nations Development Program (UNDP).',
    ))
    expect(got[0]).toMatchObject({ acronym: 'UNDP', status: 'defined', definition: 'United Nations Development Programme',
      definedAt: { paragraph: 0, start: 0, end: 43 }, uses: 2 })
    expect(analyseAcronyms(paras('SDGs (the Sustainable Development Goals) matter.'))[0]?.definition).toBe('Sustainable Development Goals')
  })

  it('matches joining words, hyphenated words and lower-case letters inside words', () => {
    const got = byName(analyseAcronyms(paras(
      'The Conference of the Parties (COP) met at COP28 with the G20.',
      'The Department for International Development (DFID) and the Bank of England (BoE) spoke.',
      'Nature-based Solutions (NbS) were funded by the African Development Bank (AfDB) and the Kreditanstalt f\u00FCr Wiederaufbau (KfW).',
      'The Organisation for Economic Co-operation and Development (OECD) agreed.',
      'Monitoring, Reporting and Verification (MRV) is required.',
    )))
    expect(got.COP?.definition).toBe('Conference of the Parties')
    expect(got.DFID?.definition).toBe('Department for International Development')
    expect(got.BoE?.definition).toBe('Bank of England')
    expect(got.NbS?.definition).toBe('Nature-based Solutions')
    expect(got.AfDB?.definition).toBe('African Development Bank')
    expect(got.KfW?.definition).toBe('Kreditanstalt f\u00FCr Wiederaufbau')
    expect(got.OECD?.definition).toBe('Organisation for Economic Co-operation and Development')
    expect(got.MRV?.definition).toBe('Monitoring, Reporting and Verification')
    expect(got.COP28?.status).toBe('undefined')
    expect(got.G20?.status).toBe('undefined')
  })

  it('does not take words from before a full stop or initials that do not match', () => {
    const got = byName(analyseAcronyms(paras(
      'Funding rose in 2020. Green Climate Fund (GCF) money followed.',
      'The local partner (LPX) signed.',
    )))
    expect(got.GCF?.definition).toBe('Green Climate Fund')
    expect(got.LPX).toMatchObject({ status: 'undefined' })
    expect(got.LPX?.definition).toBeUndefined()
  })

  it('counts headings as uses but never as definitions', () => {
    const doc: AcronymInputParagraph[] = [
      { text: 'UNDP results', isHeading: true },
      { text: 'The United Nations Development Programme (UNDP) reported.' },
      { text: 'Net Zero Cities (NZC)', isHeading: true },
      { text: 'NZC grew.' },
    ]
    const got = byName(analyseAcronyms(doc))
    expect(got.UNDP).toMatchObject({ status: 'defined-late', firstUse: { paragraph: 0, start: 0, end: 4 } })
    expect(got.NZC).toMatchObject({ status: 'undefined', uses: 2 })
    expect(got.NZC?.definition).toBeUndefined()
  })

  it('treats a two-column abbreviations table as definitions', () => {
    const doc: AcronymInputParagraph[] = [
      { text: 'IFC', inTable: true },
      { text: 'International Finance Corporation', inTable: true },
      { text: 'The IFC invests.' },
    ]
    expect(analyseAcronyms(doc)).toEqual([{
      acronym: 'IFC', status: 'defined', definition: 'International Finance Corporation', uses: 2,
      firstUse: { paragraph: 0, start: 0, end: 3 }, definedAt: { paragraph: 0, start: 0, end: 3 },
    }])
  })

  it('marks a small allowlist as common', () => {
    const got = analyseAcronyms(paras('The UK, US, EU and UN joined BwB in a UN-led review.'))
    expect(summary(got)).toEqual([
      ['BwB', 'common', undefined], ['EU', 'common', undefined], ['UK', 'common', undefined],
      ['UN', 'common', undefined], ['US', 'common', undefined],
    ])
    expect(byName(got).UN?.uses).toBe(2)
  })

  it('skips Roman numerals, currency codes, codes, formulas, URLs and all-caps phrases', () => {
    const got = analyseAcronyms(paras(
      'Annex IV, Section II, Part VIII and Chapter XI.',
      'USD 5m, GBP 2m, EUR, KES and amounts in JPY. SDG 5m in Sudan; COP 5,000 in Colombia.',
      'Deliverables D7.4, WP3.1 and T2.1; quarters Q1 and Q2; paper A4.',
      'Emissions of CO2, CO2e, CH4 and N2O.',
      'https://bwb.earth/ABC `XYZ` {icon:LEAF} [link](https://x.org/QRS) mail@UNEP.org',
      'EXECUTIVE SUMMARY',
      'ANNEX I: TERMS OF REFERENCE',
      'udp, Undp, iPhone, McKinsey and LinkedIn are not acronyms; ABCDEFGHI is too long.',
    ))
    expect(got).toEqual([])
  })

  it('keeps acronyms that are also currency codes when they are used as acronyms', () => {
    const got = byName(analyseAcronyms(paras(
      'Progress on SDG 7 and SDG 13 was reviewed at COP 28; the project IRR is 12 per cent under SSP2 and SSP.',
    )))
    expect(Object.keys(got)).toEqual(['COP', 'IRR', 'SDG', 'SSP', 'SSP2'])
    expect(got.SDG?.uses).toBe(2)
  })

  it('keeps acronyms in lists, possessives, plurals and hyphenated words', () => {
    const got = byName(analyseAcronyms(paras(
      "The AfDB, EIB, IFC and KfW joined; UNDP's report covers SMEs, NGOs, SDG-aligned bonds, the EU ETS and non-OECD states.",
    )))
    expect(Object.keys(got)).toEqual(['AfDB', 'EIB', 'ETS', 'EU', 'IFC', 'KfW', 'NGO', 'OECD', 'SDG', 'SME', 'UNDP'])
    expect(got.SME?.firstUse).toEqual({ paragraph: 0, start: 56, end: 60 })
  })

  it('returns nothing for an empty document', () => {
    expect(analyseAcronyms([])).toEqual([])
    expect(analyseAcronyms(paras('', '   '))).toEqual([])
  })
})
