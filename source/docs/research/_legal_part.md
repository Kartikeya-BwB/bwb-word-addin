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
