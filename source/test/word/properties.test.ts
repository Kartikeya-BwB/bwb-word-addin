import { afterEach, describe, expect, it } from 'vitest'
import { documentFileName, readChecklist, readDocInfo, writeChecklist, writeDocInfo } from '../../src/word/properties.ts'
import { createModel } from '../fakeWord/model.ts'
import { installFakeWord, uninstallFakeWord } from '../fakeWord/install.ts'

afterEach(uninstallFakeWord)

describe('document details', () => {
  it('writes core and BwB custom properties and reads them back', async () => {
    const model = createModel()
    installFakeWord(model)
    await writeDocInfo({
      title: 'Blended finance for adaptation', author: 'Bankers without Boundaries', docType: 'final-report',
      project: 'NZC', deliverable: 'D3.2', version: '04', funder: 'eu-grant', clientTemplate: true,
    })
    expect(model.custom.get('BwB document type')).toBe('final-report')
    expect(model.custom.get('BwB client template')).toBe('yes')
    const info = await readDocInfo()
    expect(info).toMatchObject({
      title: 'Blended finance for adaptation', author: 'Bankers without Boundaries', docType: 'final-report',
      project: 'NZC', deliverable: 'D3.2', version: '04', funder: 'eu-grant', clientTemplate: true,
    })
  })

  it('removes a custom property when its value is emptied, and ignores unknown funder values', async () => {
    const model = createModel()
    model.custom.set('BwB project', 'NZC')
    model.custom.set('BwB funder', 'somebody')
    installFakeWord(model)
    await writeDocInfo({ project: '', clientTemplate: false, deliverable: '' })
    expect(model.custom.has('BwB project')).toBe(false)
    expect((await readDocInfo()).funder).toBe('')
  })

  it('reads the file name from the document URL', () => {
    installFakeWord(createModel([], { url: 'https://bwb.sharepoint.com/sites/x/Shared%20Documents/20261013_NZC_D3.2_Report_v04.docx?web=1' }))
    expect(documentFileName()).toBe('20261013_NZC_D3.2_Report_v04.docx')
    uninstallFakeWord()
    installFakeWord(createModel())
    expect(documentFileName()).toBeNull()
  })

  it('keeps checklist ticks in the document settings', async () => {
    const model = createModel()
    installFakeWord(model)
    expect(readChecklist()).toEqual({})
    await writeChecklist({ objectives: true, facts: false })
    expect(readChecklist()).toEqual({ objectives: true, facts: false })
  })
})
