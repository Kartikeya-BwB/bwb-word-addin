// Load-and-sync semantics of Office.js: a property is readable only after load() and a later sync().
// Reading it earlier throws, which is the mistake the fake exists to catch.

export class Ctx {
  private pending: Array<() => void> = []
  syncs = 0

  queue(fn: () => void): void {
    this.pending.push(fn)
  }

  async sync(): Promise<void> {
    const run = this.pending
    this.pending = []
    run.forEach((fn) => fn())
    this.syncs += 1
  }
}

export class PropertyNotLoadedError extends Error {
  code = 'PropertyNotLoaded'
}

export abstract class FakeObject {
  protected loaded = new Set<string>()
  readonly context: Ctx

  constructor(context: Ctx) {
    this.context = context
  }

  load(spec?: string | string[] | Record<string, unknown>): this {
    const props = typeof spec === 'string' ? spec.split(',') : Array.isArray(spec) ? spec : spec ? Object.keys(spec) : ['*']
    this.context.queue(() => {
      for (const raw of props) this.markLoaded(raw.trim())
    })
    return this
  }

  /** Marks a property, or a nested path such as font/name or items/text, as loaded. */
  markLoaded(path: string): void {
    const [head, ...rest] = path.split('/')
    if (!head) return
    this.loaded.add(head)
    this.onLoaded(head, rest.join('/'))
  }

  protected onLoaded(_head: string, _rest: string): void {}

  protected need(prop: string): void {
    if (!this.loaded.has(prop) && !this.loaded.has('*')) {
      throw new PropertyNotLoadedError(`${this.constructor.name}.${prop} read before load() and sync()`)
    }
  }
}

export abstract class FakeCollection<T extends FakeObject> extends FakeObject {
  private cached: T[] | null = null

  protected abstract make(): T[]

  get items(): T[] {
    this.need('items')
    return this.cached ?? []
  }

  protected override onLoaded(head: string, rest: string): void {
    if (head !== 'items') return
    this.cached ??= this.make()
    if (rest) this.cached.forEach((item) => item.markLoaded(rest))
  }
}

/** An Office.js ClientResult: its value is filled in at the next sync. */
export function clientResult<T>(context: Ctx, initial: T, compute: () => T): { value: T } {
  const result = { value: initial }
  context.queue(() => {
    result.value = compute()
  })
  return result
}
