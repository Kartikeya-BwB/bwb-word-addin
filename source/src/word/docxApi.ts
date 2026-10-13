// The one place the Word layer imports the document builder from, so its layout can change without churn.

export { buildBlockOoxml, buildDocx } from '../docx/index.ts'
export type { Block, Content, ContentMeta, IconRenderer } from '../docx/index.ts'
