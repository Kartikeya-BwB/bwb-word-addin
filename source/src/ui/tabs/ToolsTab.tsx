import { Navigator } from '../tools/Navigator.tsx'
import { PictureTools } from '../tools/PictureTools.tsx'
import { ReviewTools } from '../tools/ReviewTools.tsx'
import { TableTools } from '../tools/TableTools.tsx'
import { TextTools } from '../tools/TextTools.tsx'
import '../tools/tools.css'

/**
 * Word power tools: pictures, tables, text, the navigator and review. The ribbon's Tools group runs the common ones
 * at once; this tab has them all, with sizes and options. The ribbon opens it at #tools, or #tools/navigator.
 */
export function ToolsTab() {
  return (
    <>
      <PictureTools />
      <TableTools />
      <TextTools />
      <Navigator />
      <ReviewTools />
    </>
  )
}
