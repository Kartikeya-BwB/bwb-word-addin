import { DiagramSection } from './DiagramSection.tsx'
import { OrgChartSection } from './OrgChartSection.tsx'
import { PictureFrameSection } from './PictureFrameSection.tsx'
import { TableToolsSection } from './TableToolsSection.tsx'
import './tools.css'

/** The slide power tools for the Insert tab: tables, org charts, diagrams and picture into frame. */
export function SlideToolsPanel() {
  return (
    <>
      <TableToolsSection />
      <OrgChartSection />
      <DiagramSection />
      <PictureFrameSection />
    </>
  )
}
