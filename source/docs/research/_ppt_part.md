### 2.1 Tools surveyed

| Tool | Pricing | Headline features | Sources |
|---|---|---|---|
| think-cell | Annual per-user subscription. Essentials from $8.90 and AI Suite from $46.54 a month (vendor FAQ); Suite $30.90 (PricingSaaS, August 2026). Free for core nonprofit use | 40+ chart types; Excel links; CAGR and difference arrows; agenda; Scale; clean up and send; proofing language; library with map templates | [features](https://www.think-cell.com/en/product/features), [FAQ](https://www.think-cell.com/en/product/benefits), [order](https://www.think-cell.com/en/order/new) |
| Macabacus | $200 Basic, $360 Professional per user a year; Enterprise on quote | Excel links (Professional); Deck Check; TurboShapes; MasterShapes; Styles; painter with chosen properties; Conform Adjustments; Track Changes (Enterprise) | [pricing](https://macabacus.com/pricing), [shape tools](https://help.macabacus.com/article/811-shape-tools) |
| UpSlide | Quote only; reportedly 5 licences minimum plus a set-up fee | Excel links; library; Slide Check; Smart Align; Swap Text | [formatting](https://upslide.com/features/excel-powerpoint-formatting/), [FitGap](https://us.fitgap.com/products/upslide) |
| Efficient Elements | €99 + VAT per user a year; fewer features on Mac | Agenda with time slots; elements (Gantt, flags, maps driven by Excel); Master Wizard; text snippets | [FAQ](https://www.efficient-elements.com/faq/), [product](https://www.efficient-elements.com/powerpoint/) |
| iSlide | Free tier; memberships about $30 to $120 a year (third parties) | Premium diagram and template libraries; one-click font, colour and layout unification | [pricing](https://www.islide.cc/pricing) |
| Office Timeline (now Lucen) | Free; Lite, Plus and Expert under $9, $17 and $21 per user a month, billed annually; "Pro+" is a legacy name; the desktop add-in is Windows only | Timelines as native shapes; swimlanes (Plus); Excel import; MS Project and Jira sync (Expert) | [pricing](https://www.officetimeline.com/pricing), [editions](https://support.officetimeline.com/difference-between-office-timeline-add-in-editions) |
| Mekko Graphics | $399 per user a year; volume $249 down to $179 | 30+ types: Marimekko, bar-mekko, cascade, bubble, Gantt | [buy](https://www.mekkographics.com/buy-now/) |
| Templafy | Free for one seat; Team $40 per user a month; enterprise on quote | Template library; smart fields; Validator; Design Converter | [pricing](https://www.templafy.com/home/pricing/), [converter](https://support.templafy.com/hc/en-us/articles/9981925474205-Design-Converter) |
| Slidewise (Neuxpower) | $79 a year; Windows | Fonts, off-brand colours, broken links; deck-wide recolour; image audit | [product](https://neuxpower.com/slidewise) |
| NXPowerLite (Neuxpower) | Perpetual, about $59 a user | Shrinks images, strips metadata, flattens embedded Excel | [product](https://neuxpower.com/nxpowerlite-desktop) |
| empower (empowersuite.com) | From about $11.99 per user a month (Capterra) | 40+ chart types; converts native charts; Excel links; Design Check; agenda | [charts](https://www.empowersuite.com/en/features/empower-ppt-charts), [Capterra](https://www.capterra.com/p/140317/empower/pricing/) |
| FlexDeck | Not found | No such product turned up in two searches; the closest names were FlexiPrez and ProDeck. Please send a URL | |
| PPT Productivity | $59, $119 or $179 (Power Tools) per seat a year | 200+ tools; MBB-style shortcuts; several format painters; slide library | [pricing](https://pptproductivity.com/pricing) |
| Pickit | Business, Business Plus or Enterprise; prices not published | Brand image library, 1M+ licensed stock images, slide library | [pricing](https://www.pickit.com/pricing) |
| SlideUpLift | Free tier; $149.99 a year for unlimited use | 40,000+ templates in an add-in | [checkout](https://slideuplift.com/checkout/?level=83) |
| Power-user (added) | About $8.75 a month (a competitor's figure) | Sankey; 350 maps coloured from Excel | [maps](https://support.powerusersoftware.com/support/solutions/articles/80001021388-library-data-maps), [Deckary](https://deckary.com/blog/powerpoint-add-ins-consultants) |
| Org charts | Visio Data Visualizer is free with a Microsoft 365 work or school account; OrgChart Now from $295 a year | Org charts from Excel data; export to PowerPoint; SmartArt is built into PowerPoint | [Microsoft](https://support.microsoft.com/en-au/visio/create-an-org-chart-based-on-excel-data-by-using-visio-for-the-web), [Software Advice](https://www.softwareadvice.com/za/hr/orgchart-profile/) |
| Maps | Esri needs an ArcGIS subscription; MapChart is free on the web | Live GIS maps (Esri); click-to-colour maps whose outputs are reportedly CC BY-SA (MapChart); think-cell maps are templates, not driven by data | [Esri](https://doc.arcgis.com/tr/maps-for-office), [MapChart](https://www.genbeta.com/paso-a-paso/asi-puedes-crear-mapa-personalizado-gratis-registro-mapchart) |
| BrightSlide | Free; VBA; Windows and Mac | Multi-painter and Match; guides; assign layouts; table sort; contrast checker | [features](https://www.brightcarbon.com/brightslide/brightslide-features-on-windows-and-mac/) |
| PowerPoint Labs | Free; GPL-2.0; no longer active | Sync Lab (painter of chosen properties); Crop to shape; Agenda Lab; Shapes Lab | [repo](https://github.com/PowerPointLabs/PowerPointLabs) |
| Instrumenta | Free; MIT; VBA | 270+ tools: table transpose and sum, multislide shapes, mail merge, proofing language | [repo](https://github.com/iappyx/Instrumenta) |
| ProDeck (added) | Free; MIT; VBA | Footnotes; select by fill; merge text; proofing language | [repo](https://github.com/rodrigolourencofarinha/ProDeck) |

### 2.2 Gaps: what premium tools offer that the toolkit did not

Numbers in brackets are PowerPointApi requirement sets. Value is to a consultancy making client decks and proposals. Size: S is under a day, M is one to three days and L is over three days.

| # | Gap: what it does | Offered by (tier) | Value | Office.js route (requirement set) | Size | Flag |
|---|---|---|---|---|---|---|
| 1 | **Keyboard shortcuts** for align, same size, painter and stickers | PPT Productivity, think-cell, UpSlide | H | KeyboardShortcuts 1.1 and `Office.actions.associate`: web, Windows 2601+, Mac 16.105.2+. Needs SharedRuntime 1.1 | M | None |
| 2 | **Selective format painter**: copies the properties you choose (position, size, fill, line, font, margins, autofit, rotation, adjustments) to shapes on any slide; presets act as named styles | Macabacus, PPT Productivity (Power Tools); free in BrightSlide and PowerPoint Labs | H | `left/top/width/height`, `ShapeFill`, `ShapeLineFormat`, `ShapeFont`, `TextFrame` margins and `autoSizeSetting` (1.4); `rotation`, `Adjustments.get/set` (1.10). Gradients, shadows and spacing cannot be copied | S | Name |
| 3 | **Edit a chart in place**, and give selected charts the same scale | think-cell, Mekko Graphics, empower | H | Spec in `Shape.tags` (1.3) or `Shape.customXmlParts` (1.7); `getSelectedShapes` (1.5); re-insert with `setSelectedDataAsync` XmlSvg (ImageCoercion 1.2) at the old frame; `Shape.delete` (1.3); `setZOrder` (1.8) | M | None |
| 4 | **Chart annotations**: CAGR arrows, level and total difference arrows, average line, totals, axis breaks | think-cell, empower | H | Drawn in the SVG renderer; no new API | M | Look |
| 5 | **More chart types**: column and line combination, labelled scatter and bubble, area; later Sankey and treemap | think-cell, Mekko Graphics, empower, Power-user (Sankey) | H (Sankey M) | SVG through XmlSvg (ImageCoercion 1.2) | S each; M for scatter | None |
| 6 | **Native charts**: embedded data clients can edit; restyle an existing chart | UpSlide, Macabacus (linked Excel charts); empower converts charts | M | There is no chart API (`ShapeType` "Chart" can only be read). Insert a generated one-slide PPTX with `insertSlidesFromBase64` (1.2), which lands as a new slide; read `c:numCache` from `Slide.exportAsBase64` (1.8) | M | None |
| 7 | **Refresh from Excel**: charts and tables re-read their range | think-cell, Macabacus (Professional), UpSlide, Mekko Graphics | M | There is no link between apps. The user re-picks the .xlsx, and the range is kept in tags (1.3) | L | None |
| 8 | **Data-driven maps**: countries shaded by value or category, with a legend | Power-user, Efficient Elements, Esri (ArcGIS subscription) | H | SVG through XmlSvg, because there is no freeform API; `Graphic.convertToShape()` is preview only | M | MapChart; borders |
| 9 | **Org chart from a pasted list** that lays itself out | Visio Data Visualizer, OrgChart Now | H | `addGeometricShape`, `addTextBox`, `addLine(ConnectorType.elbow)` (1.4), `addGroup` (1.8). Lines cannot attach to shapes or take arrowheads | M | None |
| 10 | **Diagrams**: pyramid, funnel, cycle, Venn, hub and spoke | iSlide (Premium), Efficient Elements, SlideUpLift | M | `addGeometricShape` with `trapezoid`, `blockArc`, `funnel` (1.4); `Adjustments.set`, `rotation` (1.10) | S each | Artwork |
| 11 | **Numbered footnotes**: marker at the cursor, footnote block, renumbering | Macabacus (footnote box), ProDeck, Instrumenta | H | `getSelectedTextRange` (1.5), `TextRange.getSubstring` (1.4), `ShapeFont.superscript` (1.8), tags (1.3) | M | None |
| 12 | **Text tools**: merge boxes, split them, swap text | UpSlide, PPT Productivity, Instrumenta | M | `TextRange.text`, `addTextBox` (1.4). Mixed formatting is flattened | S | None |
| 13 | **Table tools**: transpose, sort, total row, keep widths, Excel range to table | Instrumenta, BrightSlide (free), PPT Productivity | H | `Table.values`, `TableCell.text` (1.8); `rows/columns.add`, `deleteRows`, settable `TableColumn.width` and `TableRow.height`, cell formats (1.9) | M | None |
| 14 | **Ghost deck**: one slide per pasted action title | Instrumenta, Ampler | H | `slides.add({layoutId})` (1.3), `Slide.moveTo`, `placeholderFormat.type` (1.8) | S | None |
| 15 | **Deck fields**: client, project, date and version entered once and filled in everywhere | Templafy, Macabacus, Instrumenta (mail merge) | H | `properties.customProperties` (1.7), tags (1.3), footer placeholders through `PlaceholderFormat.type` (1.8). Header and Footer cannot be switched on | M | Name |
| 16 | **Restyle a deck that is not on the BwB template**: map its slides to BwB layouts | Templafy (enterprise), BrightSlide (free) | H | `insertSlidesFromBase64` with `UseDestinationTheme` (1.2), `Slide.applyLayout` (1.8), `ThemeColorScheme.setThemeColor` (1.10). Theme fonts cannot be set | M | Name |
| 17 | **Clean copy to send**: notes, comments, hidden slides and metadata removed, images shrunk | think-cell, Instrumenta, NXPowerLite, Slidewise | H | There are no APIs for these. Instead, `Document.getFileAsync(Compressed)`, edit the OOXML in the pane, then `PowerPoint.createPresentation` (1.1) | M | None |
| 18 | **Compare versions** with an earlier file | Macabacus (Enterprise) | M | `getFileAsync` and a file picker; `setSelectedSlides` (1.5) | M | None |
| 19 | **Deck check additions**: WCAG contrast, reading order, hidden shapes, notes left in, internal links | BrightSlide, Slidewise, Macabacus, UpSlide | M | `ShapeFont.color`, `ShapeFill.foregroundColor` (1.4), `Slide.background`, `Shape.visible` (1.10), `zOrderPosition` (1.8), `Slide.hyperlinks` (1.6) | S each | Name |
| 20 | **Anonymise** credential slides | think-cell, Instrumenta | M | `exportAsBase64Presentation` (1.10), replace the `a:t` text, `createPresentation` (1.1) | S | None |
| 21 | **Selection helpers**: select similar, match corner rounding | UpSlide, Macabacus, ProDeck | M | `Slide.setSelectedShapes` (1.5), `Adjustments` (1.10). An autoshape's type cannot be read | S | None |
| 22 | **Scale** a selection with its fonts, margins and line weights | think-cell | M | Geometry and `lineFormat.weight` (1.4); `ShapeFont.size` read through `getSubstring` when sizes are mixed | M | None |
| 23 | **Picture into frame**, cropped to fill | PowerPoint Labs, Instrumenta | M | There is no crop API: crop on a canvas, then `ShapeFill.setImage` (1.8) | S | None |
| 24 | **Personal slide library** | think-cell, Macabacus, PPT Productivity (Power Tools) | M | `Slide.exportAsBase64` and `getImageAsBase64` (1.8), IndexedDB, `insertSlidesFromBase64` (1.2) | M | None |
| 25 | **Proofing language** set to English (UK) | think-cell, Instrumenta, ProDeck | M | There is no API. The only route is a round trip (`exportAsBase64`, set `lang="en-GB"`, reinsert), which gives the slide a new ID | M | None |

**Flag key.**

- **Name**: the vendors' feature names are Deck Check and MasterShapes (Macabacus), Slide Check (UpSlide), Design Converter and smart fields (Templafy), Multi-painter (BrightSlide) and Sync Lab (PowerPoint Labs). **The toolkit's own "Deck check" matches Macabacus's name.** "Marimekko" is a Marimekko Oyj trademark.
- **Look**: do not copy think-cell's distinctive chart styling.
- **Artwork**: iSlide and SlideUpLift templates are licensed artwork.
- **MapChart; borders**: MapChart outputs are reportedly CC BY-SA, so do not reuse them; disputed borders need care.
- **Code**: PowerPoint Labs is GPL-2.0, so take ideas only.
- **Patents**: see Part 4.

### 2.3 The 15 worth building first

1. **Keyboard shortcuts** (M). The most-missed consulting feature. KeyboardShortcuts 1.1 now covers PowerPoint (the toolkit README said it could not; corrected).
2. **Selective format painter** (S). Used every day, and it needs only the 1.4 to 1.10 properties.
3. **Edit a chart in place** (M). Today, changing the data means rebuilding the chart.
4. **Ghost deck from an outline** (S). Makes the storyline a writing tool.
5. **Table tools** (M). Tables dominate proposals, and the Word side's logic can be ported.
6. **Deck fields** (M). Stops a stale client name surviving in a reused proposal.
7. **Combination, scatter and bubble charts** (S to M). The commonest missing types, in plain SVG.
8. **Chart annotations** (M). Arrows turn finance charts into findings.
9. **Numbered footnotes** (M). Evidence-heavy decks need sources.
10. **Clean copy to send** (M). The only dependable way to strip notes and comments.
11. **Data-driven maps** (M). Reuses the existing d3-geo pipeline for country portfolios.
12. **Org chart from a list** (M). Every proposal has a team slide.
13. **Restyle a partner deck** (M). The theme and layout APIs do most of the work.
14. **Deck check additions** (S). Cheap rules that cover what public funders ask for on accessibility.
15. **Picture into frame** (S). Pairs with the team photos in the library.

Next in line: compare versions, native charts and the selection helpers.

### 2.4 Open-source building blocks

The toolkit already ships d3-geo, topojson-client and world-atlas (ISC) and fflate (MIT).

| Library | Use | Licence |
|---|---|---|
| [d3-scale](https://github.com/d3/d3-scale), [d3-shape](https://github.com/d3/d3-shape), [d3-force](https://github.com/d3/d3-force) | Scatter, bubble and area charts | ISC |
| [d3fc label layout](https://github.com/d3fc/d3fc) | Labels that do not overlap | MIT |
| [labella.js](https://github.com/twitter/labella.js) | Milestone labels | Apache-2.0 |
| [d3-hierarchy](https://github.com/d3/d3-hierarchy) | Org charts, treemaps | ISC |
| [tidy](https://github.com/zxch3n/tidy) | Trees with boxes of different sizes | MIT |
| [dagre](https://github.com/dagrejs/dagre) | Flow diagrams | MIT |
| [d3-sankey](https://github.com/d3/d3-sankey) | Sankey diagrams | BSD-3-Clause |
| [polylabel](https://github.com/mapbox/polylabel), [simple-statistics](https://github.com/simple-statistics/simple-statistics) | Map labels; class breaks | ISC |
| [PptxGenJS](https://github.com/gitbrent/PptxGenJS) | Native charts | MIT |
| [SheetJS CE](https://github.com/SheetJS/sheetjs) | Reading .xlsx | Apache-2.0 |
| [jsdiff](https://github.com/kpdecker/jsdiff) | Comparing versions | BSD-3-Clause |
| [pica](https://github.com/nodeca/pica), [culori](https://github.com/Evercoder/culori) | Shrinking images; contrast | MIT |
| [Instrumenta](https://github.com/iappyx/Instrumenta), [ProDeck](https://github.com/rodrigolourencofarinha/ProDeck) | VBA logic to port | MIT |

**Avoid** d3-flextree (WTFPL), elkjs (EPL-2.0) and PowerPoint Labs (GPL-2.0).

### 2.5 PowerPoint API limits worth knowing

Sources: [requirement sets](https://learn.microsoft.com/javascript/api/requirement-sets/powerpoint/powerpoint-api-requirement-sets), [preview APIs](https://learn.microsoft.com/javascript/api/requirement-sets/powerpoint/powerpoint-preview-apis), [keyboard shortcuts](https://learn.microsoft.com/office/dev/add-ins/design/keyboard-shortcuts), [whole file](https://learn.microsoft.com/office/dev/add-ins/develop/get-the-whole-document-from-an-add-in-for-powerpoint-or-word).

- **Missing entirely, even in preview:**
  - Speaker notes, comments and sections, and whether a slide is hidden.
  - Animations and picture cropping.
  - Shadows and gradient fills on shapes (gradients exist only for slide backgrounds, 1.10); arrowheads.
  - Line and paragraph spacing, proofing language and theme fonts.
  - Chart contents and freeform paths.
  - Reading an autoshape's type, copying shapes, and attaching connectors to shapes.
- **No PowerPoint set is desktop-only.** The VBA tools manage notes, cropping and language only because VBA runs on the desktop.
- **Preview only:** `addPicture`, `Graphic.convertToShape`, `SlideLayout.delete` and `SlideMaster.delete`, `onSlideSelectionChanged`.
- **Version gates:**
  - 1.8 needs Windows 2504; 1.9 needs 2508; 1.10 needs 2601 or Mac 16.105.
  - Volume-licensed Office 2024 stops at 1.5, and iPad at 1.1.
  - SVG insertion (ImageCoercion 1.2) does not work on iPad.
- **Workarounds:**
  - `getFileAsync(Compressed)` reads the whole file (notes, comments, chart caches, media). Test that unsaved edits are included.
  - `createPresentation(base64)` opens a processed copy.
  - A slide round trip (`exportAsBase64` then `insertSlidesFromBase64`) gives the slide a new ID, so test comments and links.
- **Text:** shape text has no collection of runs, so rewriting `TextRange.text` flattens mixed formatting.
- **Commands:**
  - Keyboard shortcuts need a shared runtime and, on the web, do not fire while the task pane has focus.
  - Context-menu items appear only on selected text (`ContextMenuText`).
- **Masters and footers:** theme colours on masters can be written (1.10), but Header and Footer cannot be switched on. The reference does not say whether shapes can be added to masters, so test that first.
