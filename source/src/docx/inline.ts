/**
 * Inline content: text runs from the mini mark-up, inline icons, footnote references and, for the add-in, {field:Label}
 * placeholders. A port of Doc.runs, Doc.icon_xml, Doc._icon_para and Doc._footnote_ref.
 */
import type { Doc } from './builder.ts'
import { inlinePicture } from './drawing.ts'
import { checkColour, ICON_TOKEN, isIconToken, parseIconToken, stripIconTokens, suggestIconNames } from './iconPolicy.ts'
import { esc, LIME, para, pyRound, pyRoundTo, pyStrip, replaceFirst, run, splitInline, WHITE, type RunProps } from './xml.ts'

/** Run properties plus the icon size overrides the Python builder passes through **base. */
export interface RunOpts extends RunProps {
  icon_cm?: number | null
  icon_pt?: number | null
}

export const FIELD_TOKEN = '\\{field:[^{}]+\\}'
const SPLIT = new RegExp(`(${ICON_TOKEN}|${FIELD_TOKEN})`)
const FIELD_FULL = new RegExp(`^${FIELD_TOKEN}$`)
const FIELD_ALL = /\{field:([^{}]+)\}/g

/** {field:Client name} written as plain text: [Client name]. */
export function fieldsAsText(text: string): string {
  return text.replace(FIELD_ALL, (_m, label: string) => `[${pyStrip(label)}]`)
}

/**
 * An inline rich text content control holding "[Label]", for values such as the client name that the author fills in.
 * The add-in's checker looks for the tag BwBPlaceholder.
 */
export function fieldControl(doc: Doc, label: string, kw: RunProps = {}): string {
  const id = doc.nextControlId()
  return (
    `<w:sdt><w:sdtPr><w:alias w:val="${esc(label)}"/><w:tag w:val="BwBPlaceholder"/><w:id w:val="${id}"/></w:sdtPr>` +
    `<w:sdtContent>${run(`[${label}]`, kw)}</w:sdtContent></w:sdt>`
  )
}

/** run() for text that takes no mark-up in the Python builder but may hold {field:Label} placeholders. */
export function plainRuns(doc: Doc, text: string, kw: RunProps = {}): string {
  if (!text.includes('{field:')) return run(text, kw)
  let out = ''
  for (const part of text.split(new RegExp(`(${FIELD_TOKEN})`))) {
    if (!part) continue
    out += FIELD_FULL.test(part) ? fieldControl(doc, pyStrip(part.slice('{field:'.length, -1)), kw) : run(part, kw)
  }
  return out
}

/** Footnote or bracketed note text: links become "label (url)", icons are dropped, bold and italic are kept. */
export function noteRuns(text: string, kw: RunProps = {}): string {
  const plain = stripIconTokens(text).replace(/\[([^\]]+)\]\(([^)]+)\)/g, '$1 ($2)')
  let out = ''
  for (const t of splitInline(fieldsAsText(plain))) {
    if (t.kind === 'footnote') continue
    out += run(t.value, t.kind === 'bold' ? { ...kw, b: true } : t.kind === 'italic' ? { ...kw, i: true } : kw)
  }
  return out
}

export function footnoteRef(doc: Doc, text: string): string {
  const fid = doc.footnotes.length + 1
  doc.footnotes.push(text)
  return `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="${fid}"/></w:r>`
}

export interface IconPlacement {
  onNavy?: boolean
  /** Colour used when none is given, and the replacement for a colour the background does not allow. */
  def?: string | null
  fontPt?: number
  /** Baseline shift in half points; null centres the icon on the x-height of the surrounding text. */
  lower?: number | null
}

/**
 * An inline picture of a Lucide icon in an approved BwB colour: a PNG for every viewer plus the SVG for Word.
 * Returns '' (and records a warning) when the icon does not exist or cannot be drawn.
 */
export function iconXml(doc: Doc, name: string, colour: string | null | undefined, cm = 0.4, o: IconPlacement = {}): string {
  const onNavy = o.onNavy ?? false
  const def = o.def ?? null
  const fontPt = o.fontPt ?? 10
  const icons = doc.icons
  if (!icons) {
    doc.noteOnce('Icons were skipped because no icon renderer was supplied.')
    return ''
  }
  const r = icons.resolve(name)
  if (!r) {
    const sug = icons.names ? suggestIconNames(name, icons.names(), (n) => icons.resolve(n), 3) : []
    doc.noteOnce(`Icon '${name}' is not a Lucide icon` + (sug.length ? ` (did you mean ${sug.join(', ')}?)` : '') + '; skipped.')
    return ''
  }
  const { colour: c, note } = checkColour(colour || def || (onNavy ? 'white' : 'green'), onNavy, def)
  if (note) doc.noteOnce(`Icon '${r}': ${note}.`)
  const svg = icons.svg(r, c)
  if (svg === null) {
    doc.noteOnce(`Icon '${r}' could not be drawn; skipped.`)
    return ''
  }
  const base = `icon_${r}_${c}`
  const rid = doc.addIconPng(base + '.png', r, c)
  const ridSvg = doc.pkg.addMedia(base + '.svg', svg)
  const emu = pyRound(cm * 360000)
  let pic = inlinePicture(rid, emu, emu, doc.did(), `BwB icon ${r} (${c})`, `Icon: ${r.replaceAll('-', ' ')}`)
  // Word 2016 and later draw the SVG from this extension; other viewers fall back to the PNG in a:blip
  const ext =
    '<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip ' +
    `xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="${ridSvg}"/></a:ext></a:extLst>`
  pic = replaceFirst(pic, `<a:blip r:embed="${rid}"/>`, `<a:blip r:embed="${rid}">${ext}</a:blip>`)
  const lower = o.lower ?? -pyRound(Math.max(0.0, (cm / 0.0353 - 0.75 * fontPt) / 2) * 2)
  if (lower) {
    pic = replaceFirst(pic, '<w:rPr><w:noProof/></w:rPr>', `<w:rPr><w:noProof/><w:position w:val="${Math.trunc(lower)}"/></w:rPr>`)
  }
  doc.iconsUsed.add(r)
  return pic
}

/** A paragraph holding one icon, set above a title or number inside a panel. */
export function iconPara(doc: Doc, name: string, colour: string, cm: number, after = 100, keepNext = true, onNavy = false): string {
  const pic = iconXml(doc, name, colour, cm, { onNavy, lower: 0 })
  if (!pic) return ''
  return para(pic, { spacing: { before: 0, after, line: 240, lineRule: 'auto' }, keep_next: keepNext, jc: 'left' })
}

/** Runs for text with the inline mark-up (Doc.runs). */
export function inlineRuns(doc: Doc, text: unknown, base: RunOpts = {}): string {
  const { icon_cm: iconCm, icon_pt: iconPt, ...kw0 } = base
  const onNavy = doc.navy || [WHITE, LIME].includes(String(kw0.color ?? '').toUpperCase())
  const pt = iconPt || kw0.sz || 10

  const txt = (val: string, kw: RunProps): string => {
    if (!val.includes('{icon:') && !val.includes('{field:')) return run(val, kw)
    let out = ''
    for (const part of val.split(SPLIT)) {
      if (!part) continue
      if (isIconToken(part)) {
        const { name, colour } = parseIconToken(part)
        const cm = iconCm || Math.max(0.4, pyRoundTo((kw.sz || pt) * 0.0353, 2))
        out += iconXml(doc, name, colour, cm, { onNavy, def: onNavy ? 'white' : 'navy', fontPt: pt })
      } else if (FIELD_FULL.test(part)) {
        out += fieldControl(doc, pyStrip(part.slice('{field:'.length, -1)), kw)
      } else {
        out += run(part, kw)
      }
    }
    return out
  }

  let out = ''
  for (const t of splitInline(String(text))) {
    if (t.kind === 'text') out += txt(t.value, kw0)
    else if (t.kind === 'bold') out += txt(t.value, { ...kw0, b: true })
    else if (t.kind === 'italic') out += txt(t.value, { ...kw0, i: true })
    else if (t.kind === 'link') {
      const url = t.extra ?? ''
      const label = run(t.value, { rstyle: 'Hyperlink' })
      if (url.startsWith('#')) out += `<w:hyperlink w:anchor="${esc(url.slice(1))}" w:history="1">${label}</w:hyperlink>`
      else out += `<w:hyperlink r:id="${doc.pkg.addRel('hyperlink', url, true)}" w:history="1">${label}</w:hyperlink>`
    } else if (doc.fragment) {
      // insertOoxml cannot add footnotes to the open document, so the note is kept as text in square brackets
      doc.noteOnce('Footnotes cannot be inserted into an open document; their text was put in square brackets instead.')
      const before = out.endsWith('</w:t></w:r>') && !/\s<\/w:t><\/w:r>$/.test(out) ? ' ' : ''
      out += run(`${before}[`, kw0) + noteRuns(t.value, kw0) + run(']', kw0)
    } else {
      out += footnoteRef(doc, t.value)
    }
  }
  return out
}
