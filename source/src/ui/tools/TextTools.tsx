import { clearStrayFormatting, joinBrokenLines, removeEmptyParagraphs } from '../../word/tools/cleanup.ts'
import type { CaseMode } from '../../word/tools/textCase.ts'
import { EN_DASH, SYMBOLS } from '../../word/tools/textRules.ts'
import { addNonBreakingSpaces, changeSelectionCase, insertSymbol } from '../../word/tools/textTools.ts'
import { unavailable } from '../../word/tools/support.ts'
import { Disclosure, Section } from '../components.tsx'
import { useAction } from '../useAction.ts'
import { ToolButton } from './ToolButton.tsx'

const CASES: Array<{ mode: CaseMode; label: string; title: string }> = [
  { mode: 'sentence', label: 'Sentence case', title: 'Keeps acronyms such as UNDP and BwB' },
  { mode: 'title', label: 'Title Case', title: 'Leaves small words such as and, of and the in lower case' },
  { mode: 'upper', label: 'UPPER CASE', title: 'Upper case for the selection' },
  { mode: 'lower', label: 'lower case', title: 'Keeps acronyms such as UNDP and BwB' },
]

export function TextTools() {
  const { busy, run } = useAction()
  const working = busy !== null
  const reason = unavailable('core')

  const tool = (key: string, label: string, action: () => Promise<string>, title?: string) => (
    <ToolButton key={key} label={label} busy={busy === key} working={working} reason={reason} onClick={() => void run(key, action)}
      {...(title ? { title } : {})} />
  )

  const symbol = (char: string): void => void run(`symbol-${char}`, async () => {
    await insertSymbol(char)
  })

  return (
    <Section id="text" title="Text" intro="Select the text first. Each word keeps its own formatting.">
      <div className="tool-label">Change case</div>
      <div className="tool-row">
        {CASES.map((c) => tool(`case-${c.mode}`, c.label, () => changeSelectionCase(c.mode), c.title))}
      </div>
      <div className="tool-label">Tidy pasted text</div>
      <div className="tool-row">
        {tool('join', 'Join lines pasted from a PDF', joinBrokenLines, 'Joins lines broken mid-sentence, rejoins hyphenated words and removes double spaces')}
        {tool('clear', 'Clear stray formatting', clearStrayFormatting, 'Removes pasted fonts, colours and sizes; keeps bold, italic and underlining')}
      </div>
      <div className="tool-row">
        {tool('nbsp', 'Keep figures together', addNonBreakingSpaces, 'Non-breaking spaces in USD 5m, 12 per cent, 5 October 2026 and 20 MW, in the selection or the whole document')}
        {tool('empty', 'Remove empty paragraphs', removeEmptyParagraphs, 'In the selection. Paragraphs holding a picture or a break are kept')}
      </div>
      <Disclosure title="Symbols">
        <div className="symbol-grid">
          {SYMBOLS.map((s) => (
            <button key={s.char} type="button" className="symbol-btn" title={s.label} aria-label={s.label} disabled={working || Boolean(reason)}
              onClick={() => symbol(s.char)}>
              {s.char}
            </button>
          ))}
          <button type="button" className="symbol-btn wide" title="En dash, for ranges in tables only" aria-label="En dash, for ranges in tables only"
            disabled={working || Boolean(reason)} onClick={() => symbol(EN_DASH)}>
            {`${EN_DASH} in tables`}
          </button>
        </div>
        <p className="field-hint">{'In running text, write ranges with \u201Cto\u201D: 2020 to 2025. The en dash is for ranges in tables.'}</p>
      </Disclosure>
    </Section>
  )
}
