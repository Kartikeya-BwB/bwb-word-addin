# Premium Office add-ins: research for the BwB Toolkit

13 October 2026. Market and legal research by Claude (AI) for Bankers without Boundaries, at Kartikeya Aggarwal's request. Internal research, not legal advice: points marked **[Lawyer]** need a UK IP solicitor or patent attorney.

Vendor and product names appear in this research only, to say what was looked at. The toolkit itself never uses them (see [the clean-room rules](../clean-room/README.md)).

## Contents

1. [Summary and decisions](#1-summary-and-decisions)
2. [PowerPoint add-ins](#2-powerpoint-add-ins)
3. [Word and writing tools](#3-word-and-writing-tools)
4. [Legal boundaries](#4-legal-boundaries)
5. [Build plan](#5-build-plan)
6. [How this research was done](#6-how-this-research-was-done)

## 1. Summary and decisions

**The question.** Which paid add-ins do consultancies use with Word and PowerPoint? Which of their features would help BwB? And which can BwB build itself, legally, as part of its own free toolkit?

**What the premium tools cost** (indicative prices; see the tables for sources):

| Area | Tools and prices |
|---|---|
| PowerPoint toolbars and charts | think-cell from $8.90 to $46.54 a month; Macabacus $200 to $360 a year; Efficient Elements €99 a year; PPT Productivity $59 to $179 a year; Mekko Graphics $399 a year; Office Timeline up to $21 a month; UpSlide and Templafy on quote |
| Word checking and editing | PerfectIt $70 to $145 a year; WordRake $149 a year; ProWritingAid $120 a year; Grammarly about $12 a month |
| Metadata cleaning, comparison and bids | Litera, DocsCorp and Draftable mostly on quote; Loopio reported from $20,000 a year |

A typical trio (a chart add-in such as think-cell, a PowerPoint toolbar such as Macabacus and a Word consistency checker such as PerfectIt) costs roughly $380 to $1,060 per person a year at the prices above, before any bid or metadata tools.

**Where the premium value sits.**
- **PowerPoint:**
  - alignment and formatting power tools, with keyboard shortcuts;
  - charts with annotations (CAGR and difference arrows), scatter and bubble charts, and data-driven maps;
  - table tools;
  - agendas and trackers;
  - deck checks, and clean copies to send.
- **Word:**
  - consistency checks across a whole document;
  - bid aids: limits per section, compliance matrices, and names left over from an earlier bid;
  - metadata and hidden-content cleaning;
  - readability.

**What the BwB Toolkit already covered before this research.**
- PowerPoint:
  - alignment against an anchor, slide, margins or content area, plus distribute, same size, swap, tidy and snap to the grid;
  - a slide library, storyline, agenda and trackers, and stickers;
  - BwB slide elements and Harvey balls;
  - a 12-type chart builder;
  - icons, flags, maps and photos;
  - a deck check with fixes.
- Word:
  - BwB documents built from a template;
  - building blocks;
  - 23 power tools for pictures, tables and text;
  - a Pre-QA check with acronym and hyphenation consistency.

**What this round builds** (section 5 has the detail):
- Word:
  - a Bid tab: section limits, compliance matrix, name check, bid details and format rules;
  - a readability panel;
  - consistency and integrity checks: capitalisation, numbers, cross-references, table totals and lists.
- PowerPoint:
  - a format painter with saved styles;
  - table tools;
  - org charts from a list;
  - diagrams;
  - picture into frame;
  - select similar;
  - slides from an outline;
  - numbered footnotes;
  - more slide checks.
- Charts: combination, scatter, bubble and area charts; annotations; data maps; editing a chart in place.
- Both apps:
  - a "before you send" inspector, with a clean copy in PowerPoint;
  - keyboard shortcuts.

**What we deliberately do not build, and why.**

| Not built | Why |
|---|---|
| Live links from Excel to charts and tables | Office.js has no link between Office applications; it would be a product in itself |
| Suggested rewrites of wordy phrases | US patents near this feature until about 2027 (section 4.3); the toolkit only flags filler |
| Agendas that keep themselves in sync | A US patent sits near automatic synchronisation (section 4.3); the toolkit generates and updates on request |
| Native editable charts from the builder | PowerPoint and Word have no chart API; the slide library has an editable native chart instead |
| Proofing language, speaker notes and comments in PowerPoint | No API; the clean copy removes notes and comments from a copy of the file |

**The law in one paragraph.**
- What a feature does is free to build. How another product looks or reads (its code, screens, icons, help text, templates, word lists and names) is not.
- Every licence checked forbids reverse engineering, so nobody working on the toolkit installs, trials or signs up for these products. The work is based on public descriptions only.
- Vendor names never appear in the toolkit.
- Patent risk is limited to the US, and the design avoids the two areas found.
- Before the toolkit is given to anyone outside BwB or used in the US, a patent attorney should give a freedom-to-operate opinion.

## 2. PowerPoint add-ins

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

## 3. Word and writing tools

**How this was checked.** The sandbox blocked learn.microsoft.com and the vendor sites, so vendor facts come from web-search summaries of vendor and directory pages, and prices often conflict between sources: treat them as indicative. Every Office.js claim was checked in [`office-js-docs-reference`](https://github.com/OfficeDev/office-js-docs-reference), the source Microsoft generates the Learn pages from (copy of 8 October 2026). The gap list leaves out checks the toolkit already had: placeholders, highlights and small text; fake headings, skipped heading levels and long paragraphs; file name, title and author properties; comments and tracked changes left in; and a filler-phrase rule.

**Key.** WA is the `WordApi` requirement set (web, Windows, Mac and iPad). WAD is the `WordApiDesktop` set (Windows and Mac only; WAD 1.4 needs Word for Windows 2508 or Word for Mac 16.100.4, see the [requirement sets](https://learn.microsoft.com/javascript/api/requirement-sets/word/word-api-requirement-sets)). **Text** means analysis of the paragraph text the add-in already reads; **Doc** means the feature needs document APIs.

### 3.1 Tools surveyed

| Tool | Pricing (as reported) | Headline features | Sources |
|---|---|---|---|
| PerfectIt | Annual per user, reported at $70 to $145; one licence covers Windows and the web add-in (Mac) | About 30 consistency checks, editable style sheets | [checks](https://intelligentediting.com/docs/perfect-it/understanding-perfect-its-checks-index), [licence](https://intelligentediting.com/support/faqs/do-i-pay-extra-to-add-perfectit-5-or-perfectit-cloud-to-my-subscription), [Capterra](https://www.capterra.co.uk/software/208453/perfectit) |
| WordRake | $149 a year (Word), $229 with Outlook | Concision edits shown as tracked changes; acquired by BriefCatch in February 2026 | [store](https://sites.fastspring.com/wordrake/product/wordrakeforwordannuallicense), [LawNext](https://www.lawnext.com/2026/02/in-a-marriage-of-legal-editing-tools-briefcatch-has-acquired-wordrake-and-its-12-editing-patents.html) |
| Grammarly Pro, Enterprise | About $12 per member a month; Enterprise by quote | Grammar, tone, style guide, brand tone, snippets, analytics | [eesel](https://www.eesel.ai/blog/grammarly-ai-pricing), [Layer3](https://www.layer3labs.io/guides/is-grammarly-worth-it) |
| ProWritingAid | Premium $120 a year, Premium Pro $144, lifetime $399 | 20+ reports: readability, sticky sentences, overused words, consistency | [pricing](https://www.casrai.org/guides/prowritingaid-pricing), [review](https://kindlepreneur.com/prowritingaid-review/) |
| Hemingway Editor | Free on the web; Plus about $100 a year; desktop app $19.99 once | Grade level, hard sentences, adverbs, passive voice, simpler words | [Toolradar](https://toolradar.com/tools/hemingway/pricing), [review](https://diyai.io/ai-tools/text-generation/reviews/hemingway-editor-review/) |
| StyleWriter | $90 to $190 by edition | Readability index, passive index, sentence length, jargon | [Capterra](https://www.capterra.co.uk/software/179771/stylewriter4), [SCL](https://www.scl.org/1849-software-review-stylewriter-v4/) |
| Acrolinx (Markup AI since 2025) | Usage-based, from about $180 a month | Terminology, style and compliance scoring | [Slator](https://tech.slator.com/acrolinx-rebrands-markup-ai-secures-usd-27m-financing/), [Capterra](https://www.capterra.com/p/161606/Acrolinx) |
| Contract Companion, Litera Check | Quote | Defined terms, numbering repair, broken cross-references, checks as you type | [AppSource](https://marketplace.microsoft.com/en-gb/product/office/WA104380919), [sheet](https://info.litera.com/rs/046-QLX-552/images/Litera-Check_Product-Sheet_EMEA.pdf) |
| Thomson Reuters Drafting Assistant | Quote | Defined terms, cross-reference discrepancies, omissions, numbering, citations | [TR](https://uk.draftingassistant.thomsonreuters.com), [article](https://legal.thomsonreuters.com/en/insights/articles/draft-better-briefs-by-adding-ai-and-automation-to-microsoft-word) |
| Litera Compare | Quote | Redlines for Word, Excel, PowerPoint and PDF, including tables, images and OCR | [sheet](https://info.litera.com/rs/046-QLX-552/images/Litera-Compare.pdf) |
| Litera Clean (formerly Metadact) | Quote | 300+ metadata types, cleaning on send, detection of failed redactions | [sheet](https://info.litera.com/rs/046-QLX-552/images/Metadact-product-sheet.pdf), [IT Brief](https://itbrief.co.uk/story/litera-launches-clean-cloud-metadata-tool-for-firms) |
| DocsCorp cleanDocs | Quote | Removes authorship, editing time, template, comments, tracked changes and embedded Excel; recipient check | [Legal IT](https://legaltechnology.com/legal-it-newswire-product-news-docscorp-releases-the-latest-iteration-of-its-data-loss-prevention-tool-cleandocs-with-support-for-imanage-work-10/) |
| Draftable | Compare $199 to $261 per user a year; Clean from $60 | Clean (2026): risk-rated scan of 58+ metadata types, processed locally | [legal](https://draftable.com/draftable-legal), [Legal IT](https://legaltechnology.com/2026/04/01/exclusive-draftable-makes-first-move-outside-of-comparison-with-draftable-clean/) |
| Kutools for Word | $49 one-off | 100+ utilities: remove hidden text, macros and notes; batch find and replace | [order](https://www.extendoffice.com/order/kutools-for-word.html), [hidden text](https://www.extendoffice.com/product/kutools-for-word/word-remove-hidden-text.html) |
| Macabacus | $200 to $360 per user a year; Enterprise by quote | Excel-to-Word links; 200+ proofing checks (PowerPoint) | [G2](https://www.g2.com/products/corporate-finance-institute-macabacus/pricing), [blog](https://macabacus.com/blog/five-powerful-macabacus-features-for-new-users) |
| UpSlide | Quote | Excel links, content library, Smart Check (placeholders, punctuation, number format) | [support](https://support.upslide.net/hc/en-us/articles/203269765-How-to-check-your-Word-documents-for-errors) |
| Templafy | Team plan $40 per user a month; agent tiers by quote | Templates, text and asset libraries, enforced disclaimers | [pricing](https://templafy.com/pricing/) |
| Woodpecker | From $59 a month | Template automation through questionnaires and conditional logic | [Capterra](https://www.capterra.com/p/179011/Woodpecker/) |
| Loopio | Quote (reported from $20,000 a year) | Answer library, review cycles, assignment to subject experts | [Wolfia](https://wolfia.com/blog/loopio-reviews-pricing-alternatives) |
| Responsive (RFPIO) | Quote | AI requirements analysis that builds compliance matrices | [FAQ](https://www.responsive.io/faqs/compliance/does-responsive-ai-build-compliance-matrices-automatically) |
| QorusDocs | Quote | Content library, approvals, versions, Word add-in | [Capterra](https://www.capterra.co.uk/software/126738/qorus-for-proposal-management) |
| VisibleThread (also Proposal Quick Start, Docsie) | Quote, annual per user | Extracts SHALL, WILL and MUST statements and custom terms into an Excel matrix; acronyms; readability | [VT](https://visiblethread.com/compliance-matrices), [PQS](https://www.g2.com/products/proposal-quick-start/discuss), [Docsie](https://www.docsie.io/solutions/compliance-matrix/) |
| Readable | From about $8 a month | Flesch, Flesch-Kincaid, Fog, SMOG, Coleman-Liau, Spache | [tools](https://readable.com/tools/) |
| Microsoft Accessibility Checker | Free in Word | Alt text, header rows, floating objects, merged cells, heading order | [CSUN](https://w2.csun.edu/universal-design-center/documents/document-learning-tools/microsoft-office-accessibility-checker) |
| axesWord | From €329 per user a year | PDF/UA and WCAG-compliant PDFs from Word | [Capterra](https://www.capterra.com/p/10033871/axesWord/) |

### 3.2 Gaps: what premium tools offer that the toolkit did not

| # | Feature | Offered by | What it does | Value | Office.js in Word | Size | Legal |
|---|---|---|---|---|---|---|---|
| 1 | Section limits | None found; Word itself lacks it ([WPS](https://www.wps.com/office-quick-solution/excel-data-problems/word-macro-to-count-words-under-each-heading-1/)) | Words and characters with spaces under each heading, against stored limits | High | Text. `outlineLevel` (WA 1.1); live updates via `onParagraphChanged` (WA 1.6); limits stored in `Document.settings` (WA 1.4) | S | None |
| 2 | Page limits | None found ([users ask for it](https://community.responsive.io/discussion/3012/page-count-inside-of-the-project-workflow)) | Pages per section; EU portals hide pages over the limit ([template](https://www.cbe.europa.eu/system/files/2026-03/Tpl_Application%20Form%20%28Part%20B%29%20%28HE%20CBE%20JU%20IA%202026%29.docx)) | High | Doc. [`Range.pages`](https://learn.microsoft.com/javascript/api/word/word.range) (WAD 1.2); nothing equivalent on the web | M | None |
| 3 | Funder format rules | None found | Minimum font size, A4, margins of at least 15 mm, instruction pages still carrying watermarks | High | Doc. [`PageSetup`](https://learn.microsoft.com/javascript/api/word/word.pagesetup) (WAD 1.3); `Font.size` and `Section.getHeader` (WA 1.1); `Body.shapes` (WAD 1.2) | M | None |
| 4 | Compliance matrix | VisibleThread, Responsive, Proposal Quick Start, Docsie | Extracts shall, must, will and numbered items from the ToR; builds the matrix; links rows to answers; flags requirements with no answer | High | Text to extract. Doc to output: `insertTable` (WA 1.3), `insertBookmark` (WA 1.4), `customXmlParts` (WA 1.4) | M | None noted |
| 5 | Names left from an earlier bid | None found | Flags funder, client and partner names that do not belong to this bid | High | Text | S | None |
| 6 | Approved-text library | Loopio, Responsive, QorusDocs, Templafy, UpSlide | Inserts approved profiles, CVs and past projects, with review dates | High | Doc. `insertFileFromBase64` (WA 1.1; options WA 1.5); library shipped as static files | M | None |
| 7 | Bid variables | Woodpecker, Templafy | Client, funder, title and reference set once and filled in everywhere | Medium | Doc. Content controls via `getByTag` (WA 1.1) | M | None |
| 8 | Metadata clean | Litera Clean, cleanDocs, Draftable Clean | Reports and cleans properties, custom properties, comment authors' emails, personal data and revision dates | High | Doc. [`DocumentProperties`](https://learn.microsoft.com/javascript/api/word/word.documentproperties) (WA 1.3, some read-only); `customProperties.deleteAll` (WA 1.3); `Comment.authorEmail` (WA 1.4); [`removeDocumentInformation`](https://learn.microsoft.com/javascript/api/word/word.document) (WAD 1.4) | M | See Part 4 |
| 9 | Hidden content | cleanDocs, Litera Clean, Kutools | Hidden or white text, embedded Excel and OLE objects, linked files, internal links, macros | High | Doc. [`getText`](https://learn.microsoft.com/javascript/api/word/word.paragraph) with `includeHiddenText` (WA 1.7); [`Font.hidden`](https://learn.microsoft.com/javascript/api/word/word.font) (WAD 1.2); [`Field.linkFormat`](https://learn.microsoft.com/javascript/api/word/word.field), `oleFormat`, `hasVbProject` (WAD 1.4); `getHyperlinkRanges` (WA 1.3) | M | None |
| 10 | Readability | Hemingway, ProWritingAid, StyleWriter, Readable, VisibleThread | Flesch, Flesch-Kincaid, Fog and SMOG per section; hard sentences highlighted | High | Text. Show results with `Range.highlight()` (WA 1.8); cross-check `readabilityStatistics` (WAD 1.4) | S | Use generic names; the Dale-Chall and Spache word lists may be under copyright |
| 11 | Concision | WordRake, StyleWriter, Hemingway, ProWritingAid, PerfectIt | Suggests plain swaps for wordy phrases, nominalisations and EU jargon | High | Text. Fixes with `insertText` (WA 1.1) under `changeTrackingMode` (WA 1.4) | S | See Part 4; the Plain English Campaign list is under copyright |
| 12 | Style flags | Hemingway, ProWritingAid, StyleWriter | Passive voice, hedges, weasel words, intensifiers, clichés | Medium | Text | S | None (MIT word lists exist) |
| 13 | Consistency suite | PerfectIt, ProWritingAid | See 3.3 | High | Text; italics and superscripts via `Font` (WA 1.1) | S each | Do not copy style sheets |
| 14 | Cross-references | Litera Check, Drafting Assistant, PerfectIt | "Figure 7" or "Annex C" with no target, captions out of order, "Error! Reference source not found" | High | Text plus Doc: `Field.code` and `result` (WA 1.4), `type` and `updateResult` (WA 1.5); `getBookmarks` (WA 1.4) | M | None |
| 15 | Defined terms | Litera Check, Drafting Assistant | Terms such as ("the Facility") unused, used before definition, or capitalised inconsistently | Medium | Text | S | None |
| 16 | Numbers and tables | UpSlide Smart Check | Totals that do not add up, shares that do not sum to 100, million and bn written two ways | High | Doc. [`Table.values`](https://learn.microsoft.com/javascript/api/word/word.table) (WA 1.3) | M | None |
| 17 | Inline underlines | Grammarly, Litera Check | Shows findings in the text without editing it, with suggestions | Medium | Doc. `insertAnnotations` and [`Critique`](https://learn.microsoft.com/javascript/api/word/word.critique) (WA 1.7); `popupOptions` (WA 1.8) | M | None noted |
| 18 | Accessibility extras | Microsoft Accessibility Checker, axesWord | Header rows, floating images without alt text, "click here" links, colour contrast | Medium | Doc. `headerRowCount` (WA 1.3); `Shape.altTextDescription` (WAD 1.2); `Font.color` (WA 1.1) | S | None |
| 19 | Proofing settings | None found | Text marked as US English or "do not check spelling"; Word's own spelling flags | Medium | Doc. `languageId` and `hasNoProofing` (WAD 1.3); `spellingErrors` (WAD 1.4) | S | None |
| 20 | Style hygiene | Templafy, UpSlide | Styles imported from other templates; path of the attached template | Low | Doc. `getStyles` and `Style.inUse` (WA 1.5); `attachedTemplate` (WAD 1.3) | S | None |

Lower priority: version compare (Litera Compare, Draftable), where `Document.compare` (WAD 1.1) runs Word's own comparison (S); team term lists (Acrolinx, Grammarly) (S); live Excel links (Macabacus, UpSlide), for which Office.js has no link between applications (L).

### 3.3 Consistency checks worth building

From the vendor's public check index ([index](https://intelligentediting.com/docs/perfect-it/understanding-perfect-its-checks-index), [list punctuation](https://intelligentediting.com/docs/perfect-it/understanding-perfect-its-checks/bullet-list-punctuation), [titles](https://intelligentediting.com/docs/perfect-it/understanding-perfect-its-checks/missing-table-box-figure-titles.html)), restated as check types for the toolkit to implement in its own way. Each is Text unless marked Doc.

| Check type | How the toolkit would do it | Size |
|---|---|---|
| Compound words | Group "e-mail, email", "data set, dataset" and "policy maker, policymaker" by a key with spaces and hyphens removed; flag any group with two forms | S |
| Capitalisation of phrases | A phrase of one to four words capitalised mid-sentence in one place and lower case elsewhere ("Green Bonds", "Board") | S |
| Abbreviations | Used only once; spelled out again after definition; defined twice or differently; written two ways ("U.K.", "UK") | S |
| Numbers in sentences | Words in one place and numerals in another for the same range; a figure starting a sentence | S |
| Spelling variations | Two UK forms in one document (judgement and judgment), using VarCon data | S |
| Preferred spelling, phrases to avoid, similar words | A list the team can edit, with exact, near-miss and pattern entries | S |
| Accents | café and cafe; Côte d'Ivoire written two ways | S |
| Common typos | Slips that are real words (form for from, manger, trail for trial) | S |
| Contractions | "do not" and "don't" both used | S |
| Italics | The same term in italic in one place and roman in another (et al., in situ); Doc, `Font.italic` | S |
| Oxford comma | Count "A, B and C" against "A, B, and C" and flag the less common form | S |
| Superscript and subscript | Ordinals with and without superscript; CO2 against CO₂; m2; Doc | S |
| List capitalisation | The first word of each item uses the same case throughout a list | S |
| Table punctuation | Some cells end with full stops and others do not | S |
| Table and figure titles | Same form ("Figure 1:" or "Figure 1."), in order, none missing | S |

Already covered by the toolkit before this research: hyphenation, heading case, undefined abbreviations and abbreviations used before definition, unmatched brackets and quotes, list endings, comments left in, and the BwB style points. Cheap extras: a space before units (5 km), thousands separators, ellipses, and "for example" against "e.g.".

### 3.4 Sources for plain-wording suggestions

| Source | Licence | What it gives |
|---|---|---|
| [GOV.UK style guide A to Z](https://www.gov.uk/guidance/style-guide/a-to-z) | [OGL v3.0](https://www.nationalarchives.gov.uk/doc/open-government-licence/version/3/): reuse with attribution | Words to avoid (agenda, deliver, facilitate, foster, leverage) |
| US plain-language word list ([digital.gov](https://digital.gov/guides/writing-understanding/familiar-terms), [plainlanguage.gov archive](https://cybercemetery.unt.edu/oilspill/20121014022836mp_/http://www.plainlanguage.gov/howto/wordsuggestions/index.cfm)) | US federal work, so no US copyright (17 U.S.C. 105) | Hundreds of swaps (commence to begin, utilize to use) |
| [EC DGT English Style Guide](https://commission.europa.eu/system/files/2023-11/styleguide_english_dgt_en.pdf) | Reuse authorised with acknowledgement (Decision 2011/833/EU) | EU usage and terms |
| [ECA: Misused English words in EU publications](https://www.eca.europa.eu/Other%20publications/EN_TERMINOLOGY_PUBLICATION/EN_TERMINOLOGY_PUBLICATION.pdf) | Not confirmed; check the notice in the PDF | About 127 common misuses in EU English (actor, adequate, eventual) |
| [retext-simplify](https://github.com/retextjs/retext-simplify) | MIT | About 330 "use X instead of Y" patterns; the source of the list is not stated |
| [write-good](https://github.com/btford/write-good) with `too-wordy` and `no-cliches` | MIT | Wordy phrases, clichés, weasel words |
| [proselint](https://github.com/amperser/proselint) | BSD-3-Clause | Redundancy, jargon and hedging rules (Python: port the data and keep the notice) |
| Vale styles ([Microsoft](https://github.com/errata-ai/Microsoft), [Google](https://github.com/errata-ai/Google)) | MIT | Wordiness lists in YAML |
| [words/hedges](https://github.com/words/hedges), weasels, fillers | MIT | Short word lists |

Avoid: the Plain English Campaign [A to Z](https://www.plainenglish.co.uk/the-a-z-of-alternative-words) (under copyright, needs permission), the word lists in `textlint-rule-stop-words` (they [cite](https://github.com/sapegin/textlint-rule-stop-words) a commercial blog's list), and vendors' own lists.

### 3.5 The 15 worth building first

1. **Section limits panel** (S): limits are where bids most often fail. It works on every platform.
2. **Pre-send inspector and clean** (M): extends the existing send checks, with a document-inspector equivalent on desktop.
3. **Compliance matrix builder** (M): what premium bid tools charge most for; extracting modal verbs is simple and auditable.
4. **Readability per section** (S): standard formulas give the QA lead scores she can set thresholds on.
5. **Plain-wording swaps** (S): from OGL, EU and MIT lists; saves words against limits. Deferred: see Part 4.
6. **More abbreviation checks** (S): reuses the existing acronym engine.
7. **Compound words and capitalised phrases** (S): the highest-yield consistency checks.
8. **Numbers in text** (S): numerals against words, figures starting sentences, million and bn, CO2 subscripts.
9. **Cross-reference integrity** (M): broken mentions of figures and annexes are a classic QA catch.
10. **Page limits and format rules** (M): EU portals hide pages over the limit. Desktop only, so show a notice on the web.
11. **Names left from an earlier bid** (S): one wrong funder name can sink a bid.
12. **Table totals** (M): catches arithmetic slips in budgets and key figures.
13. **Remaining consistency checks** (S each): spelling variants, list capitals, figure titles, Oxford comma.
14. **Style flags** (S): passive voice, hedges and weasel words; off by default in annexes.
15. **Inline underlines** (M): writers see findings where they are, and the document does not change.

### 3.6 Open-source building blocks

| Package | Licence | Use | Size | Browser bundle |
|---|---|---|---|---|
| [`Intl.Segmenter`](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Intl/Segmenter) | Built in | Sentence and word splitting | 0 | Yes; keep a regex fallback for older webviews |
| [`syllable`](https://github.com/words/syllable) 5.0.1 | MIT | Syllable counts | 19 kB | Yes |
| [`flesch`](https://github.com/words/flesch), `flesch-kincaid`, `gunning-fog`, `smog-formula`, `coleman-liau`, `automated-readability` 2.0.1 | MIT | Readability formulas | 9 to 11 kB each | Yes |
| `text-readability` 1.1.1 | ISC | All the formulas in one package | 60 kB | Yes |
| [`retext`](https://github.com/retextjs/retext) 9 with `retext-passive`, `retext-intensify`, `retext-redundant-acronyms`, `retext-diacritics`, `retext-indefinite-article` | MIT | Passive voice, hedges, redundant acronyms, accents, a and an | 10 to 30 kB each | Yes, or copy just the word lists |
| [`write-good`](https://github.com/btford/write-good) 1.0.8 | MIT | Wordiness, clichés, weasel words, passive voice, adverbs | 42 kB plus lists | Yes, without its command-line part |
| [`sentence-splitter`](https://github.com/textlint-rule/sentence-splitter) 5.0.1 | MIT | Sentences with offsets | 215 kB | Yes |
| [`compromise`](https://github.com/spencermountain/compromise) 14.18 | MIT | Parts of speech, for nominalisations | 2.7 MB | Only if loaded on demand |
| [`textlint-rule-terminology`](https://github.com/sapegin/textlint-rule-terminology) | MIT | Term matching | 24 kB | Yes |
| [VarCon](https://github.com/en-wl/wordlist) | Permissive; keep the copyright notice | UK and US spelling variants | Data only | Yes |
| [`fflate`](https://github.com/101arrowz/fflate) 0.8.3 | MIT | Unzipping the .docx for a metadata scan | About 8 kB core | Yes (already used) |
| `number-to-words` 1.2.5 | MIT | Numbers to words | 53 kB | Yes |
| [`pdfjs-dist`](https://github.com/mozilla/pdf.js) 6.4 | Apache-2.0 | Counting PDF pages | 35 MB | Only as a worker loaded on demand |

Avoid: `sbd` (pulls in `sanitize-html`), `wink-nlp` (4.5 MB with its model), `natural` (needs Node), `text-statistics` (no licence), LanguageTool (LGPL), and `jszip` unless taken under MIT (it is dual-licensed MIT or GPL-3.0).

### 3.7 Word API limits worth knowing

- **Desktop features need gating.** WAD sets never run on the web, and WAD 1.3 and later do not run on iPad. The manifest declares no minimum requirement sets, so every button appears everywhere.
- **No page or statistics API on the web.** `Range.pages` (WAD 1.2) is the only source of page numbers. `computeStatistics` is preview only, so the toolkit counts words itself.
- **Some metadata cannot be cleared.** `lastAuthor`, `creationDate`, `lastSaveTime`, `revisionNumber` and `template` are read-only (WA 1.3). Only `removeDocumentInformation` (WAD 1.4) works like Document Inspector; its `all` and `taskpaneWebExtensions` options would also remove the add-in's own data.
- **Hidden text.** `Font.hidden` needs WAD 1.2. On the web, find hidden text by comparing `getText` with and without `includeHiddenText` (WA 1.7); removing it there means rewriting the paragraph's OOXML.
- **Whole-file access is read-only.** [`getFileAsync`](https://learn.microsoft.com/javascript/api/office/office.document) returns the .docx or a PDF in 4 MB slices (64 kB on iPad). Its reference table says Word on the web is unsupported, but the requirement-set page lists it, so test it.
- **What paragraphs miss.** Body paragraphs leave out text boxes (`Body.shapes`, WAD 1.2), headers and footers (`Section.getHeader`, WA 1.1), footnotes (`Body.footnotes`, WA 1.5) and comments (WA 1.4). `storyRanges` (WAD 1.4) covers all of these on desktop.
- **Search.** [`Body.search`](https://learn.microsoft.com/javascript/api/word/word.body) takes 255 characters at most. Run regular expressions in JavaScript and map matches back with offsets within the paragraph.
- **Anchors.** `Paragraph.uniqueLocalId` (WA 1.6) changes between sessions. To tie matrix rows to answers, use bookmarks or content controls.
- **Preview only (do not ship):** `exportAsFixedFormat`, `checkSpelling`, `countNumberedItems` and the comment events.
- **Client-only.** CORS blocks most link checks from the task pane, so flag risky targets (SharePoint links, file paths) rather than testing URLs.

## 4. Legal boundaries

### 4.0 In brief

1. Reproducing **what a feature does** is lawful in the UK and EU, and mostly in the US. Copying **how it looks or reads** (code, screens, icons, help text, templates, word lists) is not.
2. The sharpest risk is **contract**: all four EULAs checked forbid reverse engineering, and Macabacus forbids "competitive analysis". Install no trials.
3. The **patents** found are US-only, so use in the UK carries little patent risk. Two live US families sit near toolkit features: automatic agenda synchronisation, and proposed rewrites of wordy text (until about 2027). **[Lawyer]** Get a freedom-to-operate (FTO) opinion before open publication or US use.

### 4.1 Copyright: functionality versus expression

Ideas and principles, including those underlying interfaces, are unprotected ([Software Directive 2009/24/EC, Art. 1(2)](https://eur-lex.europa.eu/eli/dir/2009/24/oj); [17 U.S.C. § 102(b)](https://www.law.cornell.edu/uscode/text/17/102)).

- **SAS v World Programming ([C-406/10](https://curia.europa.eu/juris/liste.jsf?num=C-406/10), CJEU 2012):** functionality, programming language and data file formats are not protected expression, and a licensee may observe, study and test a program. A manual's wording, and its selection and arrangement of keywords and commands, can be protected.
- **SAS v WPL ([[2013] EWCA Civ 1482](https://www.bailii.org/ew/cases/EWCA/Civ/2013/1482.html)):** cloning SAS's functionality from its manuals was lawful, but WPL's own manual was held at trial to infringe where it copied SAS's wording.
- **Navitaire v easyJet ([[2004] EWHC 1725 (Ch)](https://www.5rb.com/wp-content/uploads/2013/10/Navitaire-v-Easyjet-Airline-Co-30-Jul-2004.pdf)):** no protection for "business logic", look and feel or the command set; the claimant won only on some graphical screens, as artistic works.
- **Nova v Mazooma ([[2007] EWCA Civ 219](https://www.5rb.com/wp-content/uploads/2013/10/Nova-Productions-Ltd-v-Mazooma-Games-Ltd-CA-14-Mar-2007.pdf)):** game features and rules are ideas, not expression.
- **Screens can be protected:** a GUI is a work if it is its author's own intellectual creation, but not where elements are dictated by technical function ([C-393/09 BSA](https://curia.europa.eu/juris/liste.jsf?num=C-393/09)). In [THJ v Sheridan [2023] EWCA Civ 1354](https://www.bailii.org/ew/cases/EWCA/Civ/2023/1354.html), software-generated charts and screens were original and infringed.
- **US:** a menu command hierarchy was an unprotectable "method of operation" in [Lotus v Borland, 49 F.3d 807 (1st Cir. 1995)](https://copyrightalliance.org/copyright-cases/lotus-development-v-borland), but the Supreme Court split 4 to 4, so it binds only that circuit. [Google v Oracle, 593 U.S. 1 (2021)](https://www.supremecourt.gov/opinions/20pdf/18-956_d18f.pdf) held copying Java API declarations fair use, assuming without deciding that they were copyrightable. US law on menus and interfaces is less settled.

| Element | Rule |
|---|---|
| Features, workflows, chart types, our own algorithms | Free |
| Menus and ribbon groups | Organise by BwB's workflow; never mirror a vendor's tabs, groups and order |
| Button names | Plain functional words ("Add total") are free; write all tooltips, help and messages ourselves |
| Layouts | Common conventions are fine; never rebuild a vendor's dialog or pane from screenshots |
| Icons | Never copy or redraw; use Lucide, Tabler or Fluent icons under their licences |
| Code, templates, decks, chart styles, documentation, style sheets, word lists | Never copy |

### 4.2 Licences and EULAs

| Vendor | Clause (wording from search extracts) |
|---|---|
| think-cell [Software Licence Agreement](https://www.think-cell.com/en/company/sla) | Bars reverse engineering, decompiling, disassembling its data formats and seeking the source code "except as permitted by mandatory law in accordance with Section 3.9" (which keeps German statutory rights under UrhG ss. 69d and 69e "unaffected"). Appears also to bar building a competing product (unverified). |
| Macabacus [EULA](https://macabacus.com/legal/eula), s. 5.3 | Customer shall not "(i) modify, translate, reverse engineer, decompile, disassemble, or create derivative works based on the Software except to the extent that enforcement is prohibited by applicable law" or "(viii) initiate competitive analysis or the development or provision of a competing service or product". Signing up with a work email binds the employer. |
| UpSlide [General Terms](https://upslide.com/Terms.pdf) | "The Licensee shall not modify, adapt, translate, decompile, reverse engineer, disassemble or otherwise attempt to derive UpSlide Source Code." No benchmarking clause found. |
| PerfectIt [Subscription Agreement](https://intelligentediting.com/media/70945/perfectit-subscription-agreement-tou-and-eula-june-2022.pdf), cl. 5.1.1 (June 2022; the [October 2024 version](https://static1.squarespace.com/static/63c14c194b307d521c83a4b9/t/6704f553b390fe46b9a7352e/1728378195772/PerfectIt+Service+and+Subscription+Agreement+-+Web.pdf) appears similar) | "Except with our prior written consent, you may not: ... disassemble, decompile, decrypt, hack, emulate, exploit, reverse engineer, attempt to derive the source code of, modify ..." No competition clause found. |

**In practice**

- Nobody installs a trial, signs up or accepts these licences for Toolkit work, least of all with a bwb.earth address (Macabacus would then bind BwB). Work only from public pages, manuals and videos (no log-in, no click-to-accept terms), rewritten as our own specs, copying no text or screenshots.
- UK and EU law lets a lawful user observe, study and test a program and voids contrary terms ([CDPA s. 50BA](https://www.legislation.gov.uk/ukpga/1988/48/section/50BA), [s. 296A](https://www.legislation.gov.uk/ukpga/1988/48/section/296A); Directive Arts 5(3) and 8). That protected WPL in England, but SAS still won a large US judgment on the same licence, which English courts declined to enforce (background in [[2020] EWCA Civ 599](https://www.judiciary.uk/judgments/sas-institute-inc-v-world-programming-ltd/)). The override is a defence, not a reason to become a licensee.
- **Decompiling for interoperability** (Directive Art. 6; [CDPA s. 50B](https://www.legislation.gov.uk/ukpga/1988/48/section/50B)) is narrow: lawful users only, only where interface information is not readily available, only the parts needed, and never to write a program "substantially similar in its expression". It never covers studying a competitor's internals, and we do not need it: Office.js is documented.

### 4.3 Patents

In Europe, programs and presentations of information "as such" cannot be patented ([EPC Art. 52](https://www.epo.org/en/legal/epc/2020/a52.html)). Since [Emotional Perception [2026] UKSC 3](https://www.eversheds-sutherland.com/en/united-kingdom/insights/supreme-court-overhauls-uk-patent-law-for-ai-emotional-perception-ai-v-comptroller-general-2026), the UK follows the EPO. Only technical features count towards an inventive step, and chart aesthetics and editorial rules rarely provide them. US software patents are easier to get, subject to [Alice v CLS Bank, 573 U.S. 208 (2014)](https://supreme.justia.com/cases/federal/us/573/208/). Patents are territorial: only acts in the US infringe a US patent ([35 U.S.C. § 271(a)](https://www.law.cornell.edu/uscode/text/35/271)).

What the search found, in outline:

- Every relevant patent found is a **US** patent. None has a European or UK counterpart, so use by BwB in the UK carries little patent risk.
- **Two live US families** sit near toolkit features:
  - agendas and overviews kept automatically in sync with the slides;
  - software that finds wordy text and proposes shorter rewrites (in force until about 2027).
- The toolkit is designed around both:
  - the agenda and trackers are generated and updated only when the author asks;
  - wording checks flag filler, and nothing proposes rewrites of wordy phrases.
- Patents on chart label placement exist or have recently expired. The toolkit uses simple, long-established rules (inside if it fits, otherwise outside, nudged on collision) and no published proprietary labelling algorithm.
- Nothing was found for waterfall connectors, CAGR or difference arrows, Mekko chart construction, Gantt charts from data, consistency checking or metadata cleaning. These chart forms are decades old.

The detailed patent notes (numbers, claims and estimated expiry dates) are kept by the project lead rather than published here. In the US, knowledge of a specific patent can support enhanced damages.

**[Lawyer]** Get a freedom-to-operate opinion from a patent attorney before any of these:
- using the toolkit in the US;
- giving it to anyone outside BwB;
- shipping automatic agenda synchronisation or suggested rewrites.

### 4.4 Trade marks

US registrations found: THINK-CELL ([US 3,484,171](https://trademarks.justia.com/773/62/think-77362509.html)), PERFECTIT ([US 3,974,947](https://trademarks.justia.com/775/51/perfectit-77551341.html)) and MARIMEKKO ([US 3,873,041](https://trademarks.justia.com/790/77/marimekko-79077707.html), a famous textile brand; no software class found, but marks with a reputation are protected beyond their own goods). No MEKKO GRAPHICS registration found, but insightsoftware owns the name and could rely on passing off. Treat every product name in the brief as a mark.

- No vendor name or look-alike ("-cell", "Perfect-", "-Rake", "Mekko Graphics") in the Toolkit name, ribbon, pane, file names, repository or marketing.
- Referential use ("an alternative to X") can be lawful ([Trade Marks Act 1994 s. 11(2)](https://www.legislation.gov.uk/ukpga/1994/26/section/11); EUTMR Art. 14(1)(c); US nominative fair use) but invites disputes. Leave it out.
- Five code comments in the toolkit named vendors. They were reworded to describe behaviour on 13 October 2026, and a THIRD-PARTY-NOTICES file and an About and licences section were added the same day.
- Say "for Word and PowerPoint"; use no Microsoft or Office logos ([Microsoft guidelines](https://microsoft.com/legal/intellectualproperty/trademarks)).

| Use | Avoid |
|---|---|
| Mekko chart, variable-width stacked chart | Marimekko chart, Mekko Graphics |
| Harvey balls (generic; [coined at Booz Allen](https://gobraithwaite.com/thinking/a-brief-history-of-harvey-balls/); no mark found) | Smiley-face "Harvey Ball" marks (unrelated) |
| Gantt chart, waterfall chart, CAGR arrow, difference arrow | Office Timeline; vendors' feature names |
| Consistency check, concise wording, metadata clean-up | PerfectIt, WordRake, Metadact, cleanDocs |

### 4.5 Third-party content

- **Plain English Campaign A to Z:** the PDF says "Plain English Campaign owns the copyright on this guide. You must not copy it without getting our permission first"; the [free-guides page](https://www.plainenglish.co.uk/free-guides) allows printing for your own use. Single word pairs are not protectable, but the list is likely a protected compilation. Get written permission or build our own.
- **GOV.UK style guide:** Crown copyright under [OGL v3.0](https://www.nationalarchives.gov.uk/doc/open-government-licence/version/3/). Reuse is allowed with the statement "Contains public sector information licensed under the Open Government Licence v3.0", a link to the licence and no implied endorsement.
- **US federal plain-language guidance:** no US copyright ([17 U.S.C. § 105](https://www.law.cornell.edu/uscode/text/17/105)). plainlanguage.gov was redirected to [digital.gov](https://digital.gov/guides/plain-language) in September 2025; archives exist ([Center for Plain Language](https://centerforplainlanguage.org/the-federal-plain-language-guidelines-are-missing/)). Attribution is optional; credit it anyway.
- **Linters:** [write-good](https://github.com/btford/write-good) (MIT, © Brian Ford), [proselint](https://github.com/amperser/proselint) (BSD 3-Clause, © Suchow, Pacer and Ross) and [retext-simplify](https://github.com/retextjs/retext-simplify) (MIT, © Titus Wormer; its list appears to follow the US federal list) are usable. MIT needs the copyright and licence notice in all copies; BSD also bars using the authors' names to endorse. Check each list's origin: [too-wordy](https://www.npmjs.com/package/too-wordy) points to grammarist.com, and a permissive licence cannot clear content the author did not own.

### 4.6 Clean-room guidelines (for builders)

**Permitted sources**
- Vendors' public pages, manuals and videos, read without logging in, to decide *what* to build, never *how it looks or reads*.
- BwB's Style Guide, QA Policy and Document Format skill; Microsoft's Office.js documentation.
- MIT, BSD, ISC or Apache 2.0 code with notices (ask before using GPL or AGPL); OGL and public-domain material, credited.

**Forbidden**
- Installing, trialling, signing up for or accepting the licence of any product in the brief for this work.
- Decompiling, unzipping, reading their JavaScript or VBA, watching their network traffic or opening their sample files.
- Copying or tracing text, screenshots, icons, templates, decks, chart styles, help pages, shortcut lists, word lists or style sheets.
- Asking an AI tool to reproduce a named product, or pasting competitor material into prompts.
- Using confidential knowledge from former employees of these vendors.
- Researching competitor patents further (leave that to the FTO).

**Naming**
- Generic, descriptive names only, with no vendor names or look-alikes anywhere, including comments and commit messages. Check any coined name against the UK, EU and US registers.

**Attribution**
- A THIRD-PARTY-NOTICES file and an "About and licences" pane list every third-party asset, added in the same commit: today Lucide (ISC), Tabler (MIT), flag-icons (MIT) and world-atlas (ISC; Natural Earth, public domain).

**Record of independent development**
- Per feature, a dated one-page spec in the repository listing the public sources used (URL, date, reader) and the builder.
- An exposure log: anyone who has used a listed product (for example at a client) records it and does not write that feature's spec.
- Intact git history; no competitor screenshots in the repository or chats. Keep records for at least six years after the last release.

**Ask a lawyer before**
- Publishing the Toolkit openly or giving it to anyone outside BwB.
- Shipping agenda synchronisation or concise-rewrite suggestions to US users.
- Using content with an unclear licence, or decompiling anything.
- Answering any vendor letter, or continuing if anyone has accepted one of these EULAs while working on the Toolkit.

## 5. Build plan

Eight packages of work, started on 13 October 2026. The git history shows what shipped, and each feature has a record of independent development in [docs/clean-room](../clean-room/).

| Package | App | What it adds | Main sources in this research |
|---|---|---|---|
| Bid tab | Word | Words, characters and pages per section against stored limits; a compliance matrix from the terms of reference; leftover funder and client names; bid details filled in everywhere; funder format rules | Word gaps 1 to 5 and 7 |
| Readability | Word | Reading-ease and grade scores per section, the hardest sentences, passive voice, hedges and weasel words; flags only, never rewrites | Word gaps 10 and 12 |
| Consistency and integrity checks | Word | Capitalised phrases, compound words, accents, contractions, abbreviations used once or written two ways, number styles, cross-references with no target, caption order, table totals and percentages, list and table punctuation, Oxford comma | Word gaps 13 to 16; section 3.3 |
| Before you send | Word and PowerPoint | Inspect and clean comments, tracked changes, hidden text, personal metadata, BwB notes and stickers, internal links; a clean copy of a deck without speaker notes or comments | Word gaps 8 and 9; PowerPoint gap 17 |
| Keyboard shortcuts | Word and PowerPoint | Shortcuts for alignment, sizing, ordering and styles, with a shared runtime where the app supports it | PowerPoint gap 1 |
| PowerPoint power tools | PowerPoint | Format painter with chosen properties and saved styles; table tools (format, transpose, sort, total, widths, paste from Excel); org chart from a list; picture into frame; select similar; pyramid, funnel, cycle, Venn and hub-and-spoke diagrams | PowerPoint gaps 2, 9, 10, 13, 21 and 23 |
| Charts | Word and PowerPoint | Combination, scatter, bubble and area charts; CAGR and difference arrows, average and target lines, totals and axis breaks; data-driven maps; editing a chart in place | PowerPoint gaps 3, 4, 5 and 8 |
| Slides | PowerPoint | Slides from an outline; numbered footnotes; checks for contrast, hidden shapes, internal links, duplicate titles and empty slides; "Pre-QA check" naming | PowerPoint gaps 11, 14 and 19 |

Next in line, not yet started:
- PowerPoint: deck fields (client and project names filled in everywhere), restyling a partner's deck onto BwB layouts, comparing versions, and a personal slide library.
- Word: an approved-text library, inline underlines for findings, and accessibility extras.

## 6. How this research was done

- Three AI research agents worked on 13 October 2026, one each on PowerPoint add-ins, Word and writing tools, and the legal boundaries. They used web search.
- Vendor sites and Microsoft Learn could not be opened directly from the research environment. Vendor facts therefore come from search extracts of the vendors' own pages, and prices often conflict between sources: treat them as indicative.
- Every Office.js claim was checked against the sources Microsoft builds its documentation from: the OfficeDev/office-js-docs-reference and office-js-docs-pr repositories, and the published TypeScript typings.
- EULA wording and case law come from search extracts of official pages. Check quotes against the live pages before relying on them.
- No commercial product was installed, trialled or opened during this research or the build.
