import { Section } from '../components.tsx'
import { DOC_TYPE_COUNT, NewDocument } from './NewDocument.tsx'
import { ThisDocument } from './ThisDocument.tsx'

export function StartTab() {
  return (
    <>
      <Section title="New document" weight="act"
        intro={`Start any of ${DOC_TYPE_COUNT} BwB document types: the cover, contents, numbered headings, the right structure with writing prompts, and the disclaimer.`}>
        <NewDocument />
      </Section>
      <Section title="This document" intro="Details the checks and the suggested file name use.">
        <ThisDocument />
      </Section>
    </>
  )
}
