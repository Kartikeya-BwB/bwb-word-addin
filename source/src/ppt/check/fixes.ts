// Going to a finding and applying its fix. Locations are slide and shape ids from the last snapshot, which survive
// slides moving. Every fix first checks the shape still reads as it did (its text, or its position) and skips it
// rather than guess if not. Many fixes run in a few batched syncs: find the shapes, check them, then apply.

import { UserFacingError } from '../../word/run.ts'
import { updateTrackers } from '../slides/sectionOps.ts'
import { NEEDS, shapeAt, syncEach } from '../slides/office.ts'
import { requirePpt, runPpt } from '../run.ts'
import { hasTag, tagLabel } from './shapes.ts'
import { takeDeckSnapshot } from './snapshot.ts'
import type { Box, DeckFix, DeckLocation, ShapeFix } from './types.ts'

export class StaleDeckError extends UserFacingError {
  constructor() {
    super('The slide has changed since the check ran. Run the check again.')
  }
}

export interface FixResult {
  applied: number
  skipped: number
}

type ReplaceFix = Extract<DeckFix, { kind: 'replace' }>
type TextFix = Extract<DeckFix, { expectedText: string }>

/** Shape fixes applied per sync once checked: a failure costs only its own few shapes. */
const CHUNK = 20

export function flattenFixes(fixes: DeckFix[]): DeckFix[] {
  return fixes.flatMap((f) => (f.kind === 'batch' ? flattenFixes(f.fixes) : [f]))
}

const isShapeFix = (f: DeckFix): f is ShapeFix => 'slideId' in f
const isTextFix = (f: ShapeFix): f is TextFix => 'expectedText' in f

/** Replacements in one shape, from the end of the text back so earlier offsets stay right; overlaps are skipped. */
export function orderReplacements(fixes: ReplaceFix[]): { apply: ReplaceFix[]; skipped: ReplaceFix[] } {
  const apply: ReplaceFix[] = []
  const skipped: ReplaceFix[] = []
  let limit = Number.POSITIVE_INFINITY
  for (const f of [...fixes].sort((a, b) => b.start - a.start || b.end - a.end)) {
    if (f.end > limit || f.end <= f.start) {
      skipped.push(f)
      continue
    }
    apply.push(f)
    limit = f.start
  }
  return { apply, skipped }
}

export function sameBox(a: Box, b: Box, slack = 0.5): boolean {
  return Math.abs(a.left - b.left) <= slack && Math.abs(a.top - b.top) <= slack &&
    Math.abs(a.width - b.width) <= slack && Math.abs(a.height - b.height) <= slack
}

interface Target {
  slideId: string
  path: string[]
  fixes: ShapeFix[]
  shape: PowerPoint.Shape | null
  range: PowerPoint.TextRange | null
  placeholder: PowerPoint.PlaceholderFormat | null
}

interface SpanCheck {
  target: Target
  fix: ReplaceFix
  range: PowerPoint.TextRange | null
}

/** Whether the fix still applies to the shape as it reads now. */
function stillFits(fix: ShapeFix, target: Target, shape: PowerPoint.Shape): boolean {
  const text = target.range ? target.range.text : null
  if (isTextFix(fix)) return text === fix.expectedText
  switch (fix.kind) {
    case 'align-title':
    case 'move-onto-slide':
    case 'delete-shape':
      return sameBox(shape, fix.from)
    case 'delete-empty':
      // An empty picture, chart or table placeholder may have been filled since: it must still hold nothing.
      return fix.textual ? text !== null && text.trim() === '' : target.placeholder !== null && target.placeholder.containedType === null
    default:
      return true
  }
}

/** Queues one shape's checked fixes: formatting first, then text from the end back, then any deletion. */
function queueFixes(target: Target, fixes: ShapeFix[]): void {
  const { shape, range } = target
  if (!shape) return
  for (const f of fixes) {
    switch (f.kind) {
      case 'font':
      case 'font-size':
      case 'text-colour':
        for (const s of f.spans) {
          const font = range?.getSubstring(s.start, s.length).font
          if (!font) continue
          if (f.kind === 'font') font.name = f.font
          else if (f.kind === 'font-size') font.size = f.size
          else font.color = `#${f.colour}`
        }
        break
      case 'fill-colour':
        shape.fill.setSolidColor(`#${f.colour}`)
        break
      case 'line-colour':
        shape.lineFormat.color = `#${f.colour}`
        break
      case 'align-title':
      case 'move-onto-slide':
        shape.left = f.to.left
        shape.top = f.to.top
        shape.width = f.to.width
        shape.height = f.to.height
        break
      default:
        break
    }
  }
  const replacements = fixes.filter((f): f is ReplaceFix => f.kind === 'replace')
  for (const f of orderReplacements(replacements).apply) {
    if (range) range.getSubstring(f.start, f.end - f.start).text = f.replacement
  }
  if (fixes.some((f) => f.kind === 'delete-empty' || f.kind === 'delete-shape')) shape.delete()
}

async function applyShapeFixes(fixes: ShapeFix[]): Promise<FixResult> {
  const byShape = new Map<string, Target>()
  for (const f of fixes) {
    const key = `${f.slideId}|${f.path.join('/')}`
    const target = byShape.get(key) ?? { slideId: f.slideId, path: f.path, fixes: [], shape: null, range: null, placeholder: null }
    target.fixes.push(f)
    byShape.set(key, target)
  }
  const targets = [...byShape.values()]
  return runPpt(async (context) => {
    const result: FixResult = { applied: 0, skipped: 0 }
    // 1. Find each shape and read its position, and its text where a fix depends on it.
    const missing = await syncEach(context, targets, (t) => {
      t.shape = shapeAt(context.presentation.slides.getItem(t.slideId), t.path)
      t.shape.load('left,top,width,height')
      const needsText = t.fixes.some((f) => isTextFix(f) || (f.kind === 'delete-empty' && f.textual))
      t.range = needsText ? t.shape.textFrame.textRange : null
      t.range?.load('text')
      const needsContents = t.fixes.some((f) => f.kind === 'delete-empty' && !f.textual)
      t.placeholder = needsContents ? t.shape.placeholderFormat : null
      t.placeholder?.load('containedType')
    })
    // 2. Keep the fixes whose shape still reads as it did.
    const ready = new Map<Target, ShapeFix[]>()
    for (const t of targets) {
      const shape = t.shape
      if (missing.has(t) || !shape) {
        result.skipped += t.fixes.length
        continue
      }
      const fresh = t.fixes.filter((f) => stillFits(f, t, shape))
      result.skipped += t.fixes.length - fresh.length
      if (fresh.length) ready.set(t, fresh)
    }
    // 3. Check each replacement's span still holds the text it held.
    const checks: SpanCheck[] = [...ready].flatMap(([target, list]) =>
      list.filter((f): f is ReplaceFix => f.kind === 'replace').map((fix) => ({ target, fix, range: null })))
    const unreadable = await syncEach(context, checks, (c) => {
      c.range = c.target.range ? c.target.range.getSubstring(c.fix.start, c.fix.end - c.fix.start) : null
      c.range?.load('text')
    })
    for (const c of checks) {
      if (!unreadable.has(c) && c.range && c.range.text === c.fix.match) continue
      ready.set(c.target, (ready.get(c.target) ?? []).filter((f) => f !== c.fix))
      result.skipped += 1
    }
    // 4. Apply, a few shapes per sync.
    const work = [...ready].filter(([, list]) => list.length)
    for (let i = 0; i < work.length; i += CHUNK) {
      const chunk = work.slice(i, i + CHUNK)
      let count = 0
      for (const [target, list] of chunk) {
        const overlapping = orderReplacements(list.filter((f): f is ReplaceFix => f.kind === 'replace')).skipped.length
        queueFixes(target, list)
        count += list.length - overlapping
        result.skipped += overlapping
      }
      try {
        await context.sync()
        result.applied += count
      } catch {
        result.skipped += count
      }
    }
    return result
  })
}

/** Deletes every shape carrying `tag`, or only those whose label (tag value, else text, in capitals) is `text`. */
export async function deleteTaggedShapes(tag: string, text?: string): Promise<number> {
  const snapshot = await takeDeckSnapshot({ formatting: false })
  const found = snapshot.slides
    .flatMap((slide) => slide.shapes
      .filter((s) => hasTag(s, tag) && (text === undefined || tagLabel(s, tag) === text))
      .map((shape) => ({ slideId: slide.id, path: shape.path })))
    // Shapes inside a group go before the group.
    .sort((a, b) => b.path.length - a.path.length)
  if (!found.length) return 0
  await runPpt(async (context) => {
    for (const f of found) shapeAt(context.presentation.slides.getItem(f.slideId), f.path).delete()
    await context.sync()
  })
  return found.length
}

/** Applies several fixes; returns how many worked. Shapes that changed are skipped, not guessed. */
export async function applyDeckFixes(fixes: DeckFix[]): Promise<FixResult> {
  requirePpt(NEEDS.read, 'Fixing the deck')
  const flat = flattenFixes(fixes)
  const shapeFixes = flat.filter(isShapeFix)
  const result = shapeFixes.length ? await applyShapeFixes(shapeFixes) : { applied: 0, skipped: 0 }
  const done = new Set<string>()
  for (const f of flat) {
    if (f.kind !== 'delete-tagged' && f.kind !== 'update-trackers') continue
    const key = f.kind === 'update-trackers' ? f.kind : `${f.tag}|${f.text ?? ''}`
    if (done.has(key)) continue
    done.add(key)
    if (f.kind === 'update-trackers') await updateTrackers()
    else await deleteTaggedShapes(f.tag, f.text)
    result.applied += 1
  }
  return result
}

export async function applyDeckFix(fix: DeckFix): Promise<void> {
  const { applied } = await applyDeckFixes([fix])
  if (!applied) throw new StaleDeckError()
}

/** Selects the finding's slide, then its text or shape where PowerPoint allows. */
export async function goToLocation(location: DeckLocation): Promise<void> {
  requirePpt(NEEDS.current, 'Going to a slide')
  await runPpt(async (context) => {
    try {
      context.presentation.setSelectedSlides([location.slideId])
      await context.sync()
    } catch {
      throw new StaleDeckError()
    }
    if (location.kind !== 'shape') return
    const slide = context.presentation.slides.getItem(location.slideId)
    const { start, end } = location
    if (start !== undefined && end !== undefined && end > start) {
      try {
        shapeAt(slide, location.path).textFrame.textRange.getSubstring(start, end - start).setSelected()
        await context.sync()
        return
      } catch {
        // Select the whole shape instead.
      }
    }
    try {
      slide.setSelectedShapes([location.path[0] ?? ''])
      await context.sync()
    } catch {
      throw new StaleDeckError()
    }
  })
}
