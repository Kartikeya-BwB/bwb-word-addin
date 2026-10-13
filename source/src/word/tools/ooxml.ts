// WordprocessingML checks for the text tools: taking stray run formatting out of a fragment, spotting fragments that
// are not safe to rewrite, and telling an empty paragraph from one that only looks empty.

/** Run properties that are not stray: emphasis, character styles (links, caption labels), super- and subscript. */
const KEEP = new Set([
  'rStyle', 'b', 'bCs', 'i', 'iCs', 'u', 'vertAlign', 'strike', 'dstrike', 'vanish', 'specVanish', 'webHidden', 'lang',
  'noProof', 'rtl', 'cs',
])

const RUN_PROPERTIES = /(<w:r(?:\s[^>]*)?>)(\s*)<w:rPr>([\s\S]*?)<\/w:rPr>/g

/** The child elements of a run's properties, whole, with their qualified names (w:b, w14:textFill ...). */
function childrenOf(inner: string): Array<{ name: string; xml: string }> {
  const out: Array<{ name: string; xml: string }> = []
  for (let i = inner.indexOf('<'); i >= 0; i = inner.indexOf('<', i)) {
    const open = /^<([\w.-]+:[\w.-]+)\b[^>]*?(\/?)>/.exec(inner.slice(i))
    if (!open) break
    const [tag, name = '', selfClosing] = open
    let end = i + tag.length
    if (!selfClosing) {
      const close = inner.indexOf(`</${name}>`, end)
      end = close < 0 ? inner.length : close + name.length + 3
    }
    out.push({ name, xml: inner.slice(i, end) })
    i = end
  }
  return out
}

/**
 * Removes direct run formatting (fonts, sizes, colours, highlighting, spacing, text effects) from every run in a
 * fragment, keeping bold, italic, underline, character styles and the like. The paragraph mark's own formatting is
 * left alone.
 */
export function stripRunFormatting(xml: string): { xml: string; removed: number } {
  let removed = 0
  const out = xml.replace(RUN_PROPERTIES, (_m, run: string, space: string, inner: string) => {
    const children = childrenOf(inner)
    const kept = children.filter((c) => c.name.startsWith('w:') && KEEP.has(c.name.slice(2)))
    removed += children.length - kept.length
    return kept.length ? `${run}${space}<w:rPr>${kept.map((c) => c.xml).join('')}</w:rPr>` : `${run}${space}`
  })
  return { xml: out, removed }
}

/** The document body of a package, or the text itself when it has none. */
export function bodyOf(xml: string): string {
  return /<w:body>([\s\S]*)<\/w:body>/.exec(xml)?.[1] ?? xml
}

const RISKS: Array<[RegExp, string]> = [
  [/<w:(?:ins|del|moveFrom|moveTo|rPrChange|pPrChange)\b/, 'tracked changes'],
  [/<w:comment(?:RangeStart|Reference)\b/, 'comments'],
  [/<w:(?:footnote|endnote)Reference\b/, 'footnotes'],
  [/<w:bookmarkStart\b(?![^>]*w:name="_GoBack")/, 'bookmarks'],
  [/<w:tbl\b/, 'table layout'],
]

/** What rewriting a fragment could damage (tracked changes, comments, footnotes, bookmarks, table layout), or null. */
export function fragmentRisk(xml: string): string | null {
  const body = bodyOf(xml)
  return RISKS.find(([re]) => re.test(body))?.[1] ?? null
}

const HOLDS_SOMETHING = /<w:(?:drawing|pict|object|br|sym|fldChar|fldSimple|instrText|sdt|commentRangeStart|commentReference|footnoteReference|endnoteReference|tbl|txbxContent)\b|<mc:AlternateContent\b/
const SECTION_BREAK = /<w:pPr>(?:(?!<\/w:pPr>)[\s\S])*<w:sectPr\b/
const VISIBLE_TEXT = /<w:t(?:\s[^>]*)?>[^<]*[^\s<][^<]*<\/w:t>/

/**
 * True when a paragraph that reads as empty holds nothing either: no picture, break, section break, field, content
 * control, comment, note or bookmark (other than Word's own _GoBack). Deleting it loses nothing.
 */
export function isRemovableEmpty(xml: string): boolean {
  const body = bodyOf(xml)
  if (HOLDS_SOMETHING.test(body) || SECTION_BREAK.test(body) || VISIBLE_TEXT.test(body)) return false
  return !/<w:bookmarkStart\b(?![^>]*w:name="_GoBack")/.test(body)
}
