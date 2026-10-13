import { hasPowerPoint, supportsPpt, updateMessage } from '../../../ppt/run.ts'
import { LIKENESSES, selectSimilar } from '../../../ppt/tools/similar.ts'
import { Button, Section } from '../../components.tsx'
import { useAction } from '../../useAction.ts'
import './tools.css'

/** Select similar: select one shape, then every shape on the slide that shares its fill, size, type or font. */
export function SelectSimilarSection() {
  const { busy, run } = useAction()
  const blocked = hasPowerPoint() && !supportsPpt('1.5')
  return (
    <Section id="similar" title="Select similar" intro="Select one shape, then choose what the others on the slide should share with it.">
      {blocked && <p className="tools-note">{updateMessage('Select similar')}</p>}
      <div className="tool-row">
        {LIKENESSES.map((l) => (
          <Button key={l.id} small disabled={busy !== null || blocked} onClick={() => void run(`similar-${l.id}`, () => selectSimilar(l.id))}>{l.label}</Button>
        ))}
      </div>
    </Section>
  )
}
