// The PowerPoint Insert tab: BwB slide elements and marks, charts, icons, flags and maps, photos and the logo.

import { ChartBuilder } from '../charts/ChartBuilder.tsx'
import { Section } from '../components.tsx'
import { IconBrowser } from '../library/IconBrowser.tsx'
import { ImagePicker } from '../library/ImagePicker.tsx'
import { LogoPanel } from '../library/LogoPanel.tsx'
import { ElementsPanel } from '../ppt/ElementsPanel.tsx'

export function SlideInsertTab() {
  return (
    <>
      <ElementsPanel />

      <Section id="charts" title="Charts"
        intro="Column, bar, line, waterfall, Mekko, Gantt, donut and Harvey ball tables in the BwB colours. Paste the data from Excel, then insert.">
        <ChartBuilder />
      </Section>

      <Section id="icons" title="Icons, flags and maps"
        intro="Lucide (the BwB set), the BwB sectors, Tabler, flags and maps in the brand colours. Double-click to insert on the slide.">
        <IconBrowser />
      </Section>

      <Section id="photos" title="Photos" intro="The BwB photo library: sectors, themes, markets, and the team. Double-click a photo to place it on the slide.">
        <ImagePicker />
      </Section>

      <Section title="Logo">
        <LogoPanel />
      </Section>
    </>
  )
}
