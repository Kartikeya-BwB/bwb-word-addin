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
