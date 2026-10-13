// What the PowerPoint ribbon's Agenda menu does (ExecuteFunction buttons). The names are the "fn" values in
// src/commands/ribbon-powerpoint.json; each does the same as its button in the Slides tab.

import { addTrackers, insertAgenda, markSectionHere, removeTrackers, updateTrackers } from './sectionOps.ts'

export const PPT_SLIDE_HANDLERS = {
  pptAgendaInsert: insertAgenda,
  pptTrackersAdd: addTrackers,
  pptTrackersUpdate: updateTrackers,
  pptTrackersRemove: removeTrackers,
  pptSectionMark: () => markSectionHere(),
} satisfies Record<string, () => Promise<unknown>>
