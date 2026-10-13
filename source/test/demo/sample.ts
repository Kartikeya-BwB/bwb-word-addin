// A draft report with the kinds of problems the Pre-QA check is for, used by the demo page.
import type { FakeModel } from '../fakeWord/model.ts'
import { createModel, font, para, section } from '../fakeWord/model.ts'

export function sampleModel(): FakeModel {
  const model = createModel([
    para('Blended finance for climate adaptation', 'Title', { styleBuiltIn: 'Title' }),
    para('Executive summary', 'heading 1'),
    para('This report reviews five programmes. The organization analyzed 12 MDB facilities between 2022\u20132025 and found that USD 1.2bn was mobilised.'),
    para('1. Introduction', 'heading 1'),
    para('Adaptation finance remains small \u2014 around 7 percent of climate finance flows.'),
    para('Scope and method', 'heading 3'),
    para(''),
    para(''),
    para('Figures to confirm with the client before the steering committee [insert date].', 'Normal', { font: font({ highlightColor: 'Yellow' }) }),
    para('Text pasted from an earlier proposal, still in Arial.', 'Normal', { font: font({ name: 'Arial' }) }),
    para('Key findings', 'Normal', { font: font({ bold: true }) }),
    para('The UNDP pilot shows that first-loss guarantees attract private co-investment. The Green Climate Fund (GCF) and the GCF board agreed.'),
  ])
  model.trackedChanges = 3
  model.comments = 2
  model.tables.push({ style: 'Table Grid', rows: [['Programme', 'Size (USD m)'], ['Kenya', '250'], ['Angola', '1,200']] })
  model.pictures.push({ altTextTitle: '', altTextDescription: '' })
  model.sections = [section('', 'Bankers without Boundaries', 1)]
  model.url = 'https://bwb.sharepoint.com/Shared%20Documents/Adaptation%20report%20FINAL%20-%20Copy.docx'
  model.properties = { title: '', subject: '', author: 'jsmith', company: '' }
  return model
}
