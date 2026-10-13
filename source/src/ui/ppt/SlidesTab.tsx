import { useCallback, useEffect, useMemo, useState } from 'react'
import type { ReactNode } from 'react'
import { BookmarkPlus, ClipboardCopy, Copy, ListOrdered, ListTree, RefreshCw, Trash2 } from 'lucide-react'
import { copySlidesToNewDeck } from '../../ppt/slides/copy.ts'
import type { LibrarySlide, SlideFormatting, SlideLibraryIndex } from '../../ppt/slides/library.ts'
import { FORMATTING_OPTIONS, insertLibrarySlide, loadSlideIndex, searchSlides, slideCategories } from '../../ppt/slides/library.ts'
import { NEEDS } from '../../ppt/slides/office.ts'
import { hasPowerPoint, supportsPpt, unavailableReason } from '../../ppt/run.ts'
import type { DeckOutline } from '../../ppt/slides/outline.ts'
import { readOutline } from '../../ppt/slides/outline.ts'
import {
  addTrackers, insertAgenda, markSectionHere, removeSectionMark, removeTrackers, renameSection, updateTrackers,
} from '../../ppt/slides/sectionOps.ts'
import type { Section as DeckSection } from '../../ppt/slides/sections.ts'
import { trackerText } from '../../ppt/slides/sections.ts'
import { goToSlide, insertExecutiveSummary, storylineIssues, storylineText } from '../../ppt/slides/storyline.ts'
import { errorMessage } from '../../word/run.ts'
import { Button, Notice, Section, SelectField, TextField } from '../components.tsx'
import { useNotify } from '../toast.tsx'
import { useAction } from '../useAction.ts'
import './slides.css'

const plural = (n: number, one: string, many = `${one}s`): string => `${n} ${n === 1 ? one : many}`

function SlideLibrary() {
  const { busy, run } = useAction()
  const [index, setIndex] = useState<SlideLibraryIndex | null>(null)
  const [failed, setFailed] = useState<string | null>(null)
  const [query, setQuery] = useState('')
  const [category, setCategory] = useState('')
  const [selected, setSelected] = useState<LibrarySlide | null>(null)
  const [formatting, setFormatting] = useState<SlideFormatting>('KeepSourceFormatting')

  useEffect(() => {
    let live = true
    loadSlideIndex().then((i) => live && setIndex(i), (e: unknown) => live && setFailed(errorMessage(e)))
    return () => {
      live = false
    }
  }, [])

  const results = useMemo(() => (index ? searchSlides(index, query, category) : []), [index, query, category])
  const insert = (slide: LibrarySlide): void =>
    void run('library', async () => {
      const done = await insertLibrarySlide(slide, formatting, index ?? undefined)
      return `“${slide.title}” inserted ${done.afterCurrent ? 'after this slide' : 'at the end of the deck'}. Type over its sample text.`
    })

  const option = FORMATTING_OPTIONS.find((o) => o.value === formatting)
  const tooOld = unavailableReason(NEEDS.tags, 'Inserting library slides')
  return (
    <Section id="library" title="Slide library" intro="BwB slides ready to fill in. Double-click a slide to insert it after the current one.">
      {tooOld && <Notice kind="warn">{tooOld}</Notice>}
      {failed && <Notice kind="error">{failed}</Notice>}
      {!failed && !index && <p className="muted small">Loading the slide library…</p>}
      {index && (
        <>
          <TextField label={`Search ${index.slides.length} BwB slides`} type="search" value={query} onChange={setQuery} placeholder="agenda, timeline, figures" />
          <div className="row chips" aria-label="Categories">
            {[{ id: '', label: 'All' }, ...slideCategories(index)].map((c) => (
              <button key={c.id || 'all'} type="button" className="chip" aria-pressed={category === c.id} onClick={() => setCategory(c.id)}>{c.label}</button>
            ))}
          </div>
          {results.length === 0 && <p className="muted small">No slide matches “{query}”.</p>}
          <div className="slide-grid" role="listbox" aria-label="Library slides">
            {results.map((s) => (
              <button key={s.id} type="button" role="option" className="slide-cell" aria-selected={selected?.id === s.id} title={s.description}
                onClick={() => setSelected(s)} onDoubleClick={() => insert(s)}>
                <img src={`./${s.thumb}`} alt="" loading="lazy" />
                <span>{s.title}</span>
              </button>
            ))}
          </div>
          {selected?.description && <p className="field-hint slide-description">{selected.description}</p>}
          <SelectField<SlideFormatting> label="Formatting" value={formatting} onChange={setFormatting}
            options={FORMATTING_OPTIONS.map((o) => ({ value: o.value, label: o.label }))} {...(option ? { hint: option.hint } : {})} />
          <Button variant="primary" block disabled={!selected} busy={busy === 'library'} onClick={() => selected && insert(selected)}>
            {selected ? `Insert “${selected.title}”` : 'Choose a slide'}
          </Button>
        </>
      )}
    </Section>
  )
}

function Storyline({ outline, busy, act, onRead }: Pick<Panel, 'outline' | 'busy' | 'act' | 'onRead'>) {
  const notify = useNotify()
  const canGo = supportsPpt(NEEDS.current)
  const goNote = unavailableReason(NEEDS.current, 'Going to a slide from the list')
  const copy = async (): Promise<void> => {
    if (!outline) return
    try {
      await navigator.clipboard.writeText(storylineText(outline.entries))
      notify('Storyline copied.')
    } catch {
      notify('The clipboard is not available here.', 'error')
    }
  }
  const issues = outline ? storylineIssues(outline.entries) : null
  let section: string | null = null
  return (
    <Section id="storyline" title="Storyline" intro="Every slide’s title in order, to read the argument from start to finish. Choose a title to go to its slide.">
      <Button block busy={busy === 'read'} onClick={onRead}>
        {outline ? <><RefreshCw size={13} aria-hidden /> Read the storyline again</> : <><ListTree size={13} aria-hidden /> Read the storyline</>}
      </Button>
      {issues && (issues.untitled || issues.long) ? (
        <Notice kind="warn">
          {[issues.untitled ? `${plural(issues.untitled, 'slide')} without a title` : '', issues.long ? `${plural(issues.long, 'title')} over two lines` : '']
            .filter(Boolean).join('; ')}.
        </Notice>
      ) : null}
      {outline && (
        <>
          <ol className="story-list" aria-label="Slide titles">
            {outline.entries.map((e) => {
              const heading = e.section !== null && e.section !== section ? e.section : null
              section = e.section
              return (
                <li key={e.slideId}>
                  {heading && <div className="story-section">{heading}</div>}
                  <button type="button" className="story-item" aria-current={outline.current === e.slideId || undefined} disabled={!canGo}
                    onClick={() => act(`go:${e.slideId}`, async () => { await goToSlide(e.slideId) })}>
                    <span className="story-num">{e.number}</span>
                    <span className={`story-title ${e.title ? '' : 'untitled'}`}>{e.title ?? 'No title'}</span>
                    {!e.title && <span className="badge major">No title</span>}
                    {e.long && <span className="badge minor">Over two lines</span>}
                  </button>
                </li>
              )
            })}
          </ol>
          {goNote && <p className="field-hint">{goNote}</p>}
          <div className="row">
            <Button onClick={() => void copy()}><ClipboardCopy size={13} aria-hidden /> Copy storyline</Button>
            <Button busy={busy === 'summary'} onClick={() => act('summary', insertExecutiveSummary)}>
              <ListOrdered size={13} aria-hidden /> Insert executive summary slide
            </Button>
          </div>
          <p className="field-hint">The summary lists the titles of the content slides as bullets, on a new slide after the current one.</p>
        </>
      )}
    </Section>
  )
}

function SectionRow({ section, n, busy, act }: { section: DeckSection; n: number; busy: string | null; act: Panel['act'] }) {
  const [name, setName] = useState(section.name)
  useEffect(() => setName(section.name), [section.name])
  return (
    <div className="section-row">
      <span className="section-num">{String(n).padStart(2, '0')}</span>
      <div className="grow">
        <TextField label={`Section ${n}, from slide ${section.index + 1}${section.from === 'divider' ? ' (divider)' : ''}`} value={name} onChange={setName} />
      </div>
      <div className="section-actions">
        <Button small disabled={name.trim() === section.name || !name.trim()} busy={busy === `rename:${section.slideId}`}
          onClick={() => act(`rename:${section.slideId}`, () => renameSection(section.slideId, name))}>Rename</Button>
        {section.from === 'tag' && (
          <Button small title="Remove this section" aria-label={`Remove section ${section.name}`} busy={busy === `remove:${section.slideId}`}
            onClick={() => act(`remove:${section.slideId}`, () => removeSectionMark(section.slideId))}><Trash2 size={13} aria-hidden /></Button>
        )}
      </div>
    </div>
  )
}

function Sections({ outline, busy, act }: Pick<Panel, 'outline' | 'busy' | 'act'>) {
  const sections = outline?.sections ?? []
  const sample = sections.length ? trackerText(sections, Math.min(1, sections.length - 1)) : null
  const t = outline?.trackers
  return (
    <Section id="sections" title="Sections and agenda"
      intro="Sections start at divider slides, or at any slide you mark. The agenda and the section trackers list them.">
      <Button block busy={busy === 'mark'} onClick={() => act('mark', () => markSectionHere())}>
        <BookmarkPlus size={13} aria-hidden /> Start a section at this slide
      </Button>
      <p className="field-hint">{unavailableReason(NEEDS.current, 'Starting a section at the current slide') ?? 'The section takes the slide’s title as its name. Shorten it below if it is long.'}</p>
      {outline && !sections.length && <p className="muted small">No sections yet.</p>}
      {sections.length > 0 && (
        <div className="section-list">
          {sections.map((s, i) => <SectionRow key={s.slideId} section={s} n={i + 1} busy={busy} act={act} />)}
        </div>
      )}
      <Button block disabled={!sections.length} busy={busy === 'agenda'} onClick={() => act('agenda', insertAgenda)}>
        <ListOrdered size={13} aria-hidden /> Insert agenda slide
      </Button>
      <h3>Section trackers</h3>
      {sample && (
        <p className="tracker-preview" aria-label="How a tracker looks">
          {sample.text.slice(0, sample.current.start)}<strong>{sample.text.slice(sample.current.start, sample.current.start + sample.current.length)}</strong>
          {sample.text.slice(sample.current.start + sample.current.length)}
        </p>
      )}
      {t && t.slides > 0 && (
        t.issues.length
          ? <Notice kind="warn">Trackers on {plural(t.slides, 'slide')}; out of date on {plural(t.issues.length, 'slide')}. Update them.</Notice>
          : <p className="muted small">Trackers on {plural(t.slides, 'slide')}, up to date.</p>
      )}
      <div className="row">
        <Button disabled={!sections.length} busy={busy === 'trackers-add'} onClick={() => act('trackers-add', addTrackers)}>Add</Button>
        <Button disabled={!sections.length} busy={busy === 'trackers-update'} onClick={() => act('trackers-update', updateTrackers)}>Update</Button>
        <Button busy={busy === 'trackers-remove'} onClick={() => act('trackers-remove', removeTrackers)}>Remove</Button>
      </div>
      <p className="field-hint">A row above the title of every slide from the first section on, with the current section in navy bold. Title, agenda, divider and closing slides have none.</p>
    </Section>
  )
}

function CopyToNewDeck() {
  const { busy, run } = useAction()
  const tooOld = unavailableReason(NEEDS.export, 'Copying slides to a new presentation')
  return (
    <Section id="copy" title="Copy slides to a new deck" intro="Select slides in the thumbnails on the left, then copy them into a new presentation: for a pre-read, an extract or a backup.">
      {tooOld && <Notice kind="warn">{tooOld}</Notice>}
      <Button block disabled={Boolean(tooOld)} busy={busy === 'copy'} onClick={() => void run('copy', copySlidesToNewDeck)}>
        <Copy size={13} aria-hidden /> Copy the selected slides to a new deck
      </Button>
    </Section>
  )
}

interface Panel {
  outline: DeckOutline | null
  busy: string | null
  act: (key: string, action: () => Promise<string | void>) => void
  onRead: () => void
}

/** The Slides tab: library, storyline, sections and agenda, stickers and copying slides out. */
export function SlidesTab({ stickers }: { stickers?: ReactNode }) {
  const { busy, run } = useAction()
  const [outline, setOutline] = useState<DeckOutline | null>(null)
  const refresh = useCallback(async (): Promise<DeckOutline> => {
    const next = await readOutline()
    setOutline(next)
    return next
  }, [])

  useEffect(() => {
    if (hasPowerPoint() && supportsPpt(NEEDS.read)) refresh().catch(() => undefined)
  }, [refresh])

  // Every change re-reads the deck, so the storyline and sections stay true to it.
  const act: Panel['act'] = (key, action) =>
    void run(key, async () => {
      const message = await action()
      await refresh().catch(() => undefined)
      return message
    })
  const onRead = (): void =>
    void run('read', async () => {
      const o = await refresh()
      return `${plural(o.entries.length, 'slide')} read.`
    })

  const tooOld = unavailableReason(NEEDS.read, 'The Slides tab')
  return (
    <>
      {tooOld && <Notice kind="warn">{tooOld}</Notice>}
      <SlideLibrary />
      <Storyline outline={outline} busy={busy} act={act} onRead={onRead} />
      <Sections outline={outline} busy={busy} act={act} />
      <Section id="stickers" title="Stickers" intro="Status labels such as Draft or For discussion at the top right of the selected slides.">
        {stickers ?? <p className="muted small">Use Sticker on the BwB ribbon tab. The deck check lists any left on slides.</p>}
      </Section>
      <CopyToNewDeck />
    </>
  )
}
