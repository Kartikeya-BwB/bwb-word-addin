// Office.js load-and-sync semantics for the fake PowerPoint, modelled on test/fakeWord/base.ts. A property is
// readable only after load() and a later sync(); writes and method calls are queued and applied, in order, on sync.
// An API newer than the fake's PowerPointApi version fails the sync with ApiNotFound, as PowerPoint does.

export class FakeOfficeError extends Error {
  readonly code: string
  readonly debugInfo: { code: string; message: string }

  constructor(code: string, message: string) {
    super(message)
    this.name = 'RichApi.Error'
    this.code = code
    this.debugInfo = { code, message }
  }
}

export class PropertyNotLoadedError extends FakeOfficeError {
  constructor(what: string) {
    super('PropertyNotLoaded', `${what} read before load() and sync()`)
  }
}

/** Compares requirement set versions as numbers per part, so 1.10 is newer than 1.9. */
export function compareVersions(a: string, b: string): number {
  const pa = a.split('.').map(Number)
  const pb = b.split('.').map(Number)
  for (let i = 0; i < Math.max(pa.length, pb.length); i += 1) {
    const d = (pa[i] ?? 0) - (pb[i] ?? 0)
    if (d !== 0) return d
  }
  return 0
}

export class Ctx {
  private pending: Array<() => void> = []
  syncs = 0

  constructor(private readonly version: () => string) {}

  queue(fn: () => void): void {
    this.pending.push(fn)
  }

  supports(version: string): boolean {
    return compareVersions(this.version(), version) >= 0
  }

  /** Queues an ApiNotFound failure when the fake's PowerPointApi version is older than `version`. */
  gate(version: string, what: string): void {
    if (this.supports(version)) return
    const running = this.version()
    this.queue(() => {
      throw new FakeOfficeError('ApiNotFound', `${what} needs PowerPointApi ${version}; this PowerPoint has ${running}.`)
    })
  }

  /** Applies queued operations in order. The first failure rejects the sync and drops the rest, like a batch. */
  async sync(): Promise<void> {
    const run = this.pending
    this.pending = []
    this.syncs += 1
    for (const fn of run) fn()
  }
}

export type LoadSpec = string | string[] | Record<string, unknown> | undefined

/** Turns every form load() accepts into slash paths: 'a,b/c', ['a'], { select, expand } and nested option objects. */
export function loadPaths(spec: LoadSpec): string[] {
  if (spec === undefined) return ['*']
  if (typeof spec === 'string') return spec.split(',').map((s) => s.trim()).filter(Boolean)
  if (Array.isArray(spec)) return spec.flatMap((s) => loadPaths(s))
  if (typeof spec.select === 'string' || typeof spec.expand === 'string') {
    return [...loadPaths(String(spec.select ?? '')), ...loadPaths(String(spec.expand ?? ''))]
  }
  const out: string[] = []
  for (const [key, value] of Object.entries(spec)) {
    if (key === '$all') {
      if (value) out.push('*')
    } else if (value === true) {
      out.push(key)
    } else if (value && typeof value === 'object') {
      out.push(...loadPaths(value as Record<string, unknown>).map((p) => (p === '*' ? key : `${key}/${p}`)))
    }
  }
  return out
}

export abstract class FakeObject {
  protected readonly loaded = new Set<string>()
  private synced = false
  readonly context: Ctx

  constructor(context: Ctx) {
    this.context = context
  }

  /** Requirement sets for properties newer than the object itself, such as rotation (1.10) on a shape. */
  protected versionFor(_path: string): string | undefined {
    return undefined
  }

  /** Checked when a load is applied: throw a FakeOfficeError to fail the sync (a null object, a missing frame). */
  protected checkLoad(): void {}

  load(spec?: LoadSpec): this {
    const paths = loadPaths(spec)
    for (const p of paths) {
      const version = this.versionFor(p)
      if (version) this.context.gate(version, `${this.constructor.name}.${p}`)
    }
    this.context.queue(() => {
      this.checkLoad()
      this.synced = true
      for (const p of paths) this.markLoaded(p)
    })
    return this
  }

  /** Marks a property, or a nested path such as textFrame/textRange/text, as loaded. */
  markLoaded(path: string): void {
    const [head, ...rest] = path.split('/')
    if (!head) return
    this.synced = true
    this.loaded.add(head)
    this.onLoaded(head, rest.join('/'))
  }

  /** Forwards nested paths to child objects. A navigation property loaded on its own loads all its scalars. */
  protected onLoaded(_head: string, _rest: string): void {}

  protected need(prop: string): void {
    if (!this.loaded.has(prop) && !this.loaded.has('*')) throw new PropertyNotLoadedError(`${this.constructor.name}.${prop}`)
    if (this.isNull()) {
      throw new FakeOfficeError('InvalidOperation', `${this.constructor.name}.${prop} read on a null object; check isNullObject first`)
    }
  }

  /** True for what a *OrNullObject method returns when nothing matches. Readable after any load and sync. */
  get isNullObject(): boolean {
    if (!this.synced) throw new PropertyNotLoadedError(`${this.constructor.name}.isNullObject`)
    return this.isNull()
  }

  protected isNull(): boolean {
    return false
  }

  /** Queues a write; on a null object the sync fails, as the add-in should have checked isNullObject. */
  protected write(fn: () => void): void {
    this.context.queue(() => {
      if (this.isNull()) throw new FakeOfficeError('ItemNotFound', `${this.constructor.name} is a null object`)
      fn()
    })
  }
}

/** A collection: items are listed when the load is applied, so a later load sees shapes added since. */
export abstract class FakeCollection<T extends FakeObject> extends FakeObject {
  private cachedItems: T[] | null = null
  private readonly proxies = new Map<object, T>()

  /** The data behind the items, at the time of the call. */
  protected abstract source(): object[]
  protected abstract wrap(data: object): T

  /** The same proxy for the same data, so loaded properties survive a second load. */
  protected proxy(data: object): T {
    const found = this.proxies.get(data)
    if (found) return found
    const made = this.wrap(data)
    this.proxies.set(data, made)
    return made
  }

  get items(): T[] {
    this.need('items')
    return this.cachedItems ?? []
  }

  /** Item properties newer than the collection, keyed by the property name. */
  protected itemVersions(): Record<string, string> {
    return {}
  }

  protected override versionFor(path: string): string | undefined {
    const p = path.startsWith('items/') ? path.slice(6) : path
    return this.itemVersions()[p.split('/')[0] ?? '']
  }

  // collection.load('name') loads the items' name, as in Office.js; load() and load('items') load their scalars.
  override markLoaded(path: string): void {
    const rest = path === '*' || path === 'items' ? '*' : path.startsWith('items/') ? path.slice(6) : path
    this.loaded.add('items')
    super.markLoaded('items')
    this.cachedItems = this.source().map((d) => this.proxy(d))
    for (const item of this.cachedItems) item.markLoaded(rest)
  }
}

/** What getCount() and similar methods return: the value is readable after the next sync. */
export class FakeClientResult<T> {
  private ready = false
  private result: T | undefined

  constructor(context: Ctx, compute: () => T) {
    context.queue(() => {
      this.result = compute()
      this.ready = true
    })
  }

  get value(): T {
    if (!this.ready) throw new PropertyNotLoadedError('ClientResult.value')
    return this.result as T
  }
}
