// The fake PowerPoint and Office globals: install them over a FakePresentation, run the add-in's PowerPoint code,
// inspect the model. See index.ts for the whole picture.

import { Ctx, FakeOfficeError, compareVersions } from './base.ts'
import type { FakePresentation, FakeShape } from './model.ts'
import { picture } from './model.ts'
import { activeSlide } from './queries.ts'
import { PresentationObj } from './presentation.ts'

export interface FakePowerPointHandle {
  model: FakePresentation
  /** How many times the add-in called PowerPoint.run. */
  runs(): number
  /** How many syncs those runs made, including the one run() makes at the end. */
  syncs(): number
}

export interface InstallOptions {
  /**
   * Overrides model.apiVersion: the newest PowerPointApi set this PowerPoint supports. It is written into the model,
   * so installing the same deck again without it keeps this version.
   */
  apiVersion?: string
}

/** Office enums as plain strings: PowerPoint.ShapeZOrder.bringToFront is 'BringToFront', and so on. */
const pascalEnum = (): Record<string, string> =>
  new Proxy({}, { get: (_t, key) => (typeof key === 'string' ? key.charAt(0).toUpperCase() + key.slice(1) : undefined) })

type Callback = (result: { status: string; value?: unknown; error?: { code: number; message: string } }) => void

const SUCCEEDED = 'succeeded'
const FAILED = 'failed'

/** Office.context.document.setSelectedDataAsync: records the call; a picture or SVG lands on the active slide, selected. */
function setSelectedDataAsync(model: FakePresentation, data: string, optionsOrCallback?: unknown, maybeCallback?: Callback): void {
  const options = (typeof optionsOrCallback === 'object' && optionsOrCallback !== null ? optionsOrCallback : {}) as Record<string, unknown>
  const callback = (typeof optionsOrCallback === 'function' ? optionsOrCallback : maybeCallback) as Callback | undefined
  if (model.failNextDataInsert !== null) {
    const message = model.failNextDataInsert
    model.failNextDataInsert = null
    callback?.({ status: FAILED, error: { code: 5001, message } })
    return
  }
  const coercionType = String(options.coercionType ?? 'text')
  const num = (key: string, fallback: number): number => (typeof options[key] === 'number' ? (options[key] as number) : fallback)
  const slide = activeSlide(model)
  let added: FakeShape | null = null
  if ((coercionType === 'image' || coercionType === 'xmlSvg') && slide) {
    const width = num('imageWidth', 100)
    const height = num('imageHeight', 100)
    const { slideWidth, slideHeight } = model.pageSetup
    added = picture({
      type: coercionType === 'xmlSvg' ? 'Graphic' : 'Image',
      left: num('imageLeft', (slideWidth - width) / 2),
      top: num('imageTop', (slideHeight - height) / 2),
      width,
      height,
    })
    added.name = `${coercionType === 'xmlSvg' ? 'Graphic' : 'Picture'} ${added.id}`
    slide.shapes.push(added)
    model.selection = { slideIds: [slide.id], shapeIds: [added.id], textRange: null }
  }
  model.dataInserts.push({ data, coercionType, options, shapeId: added?.id ?? null })
  callback?.({ status: SUCCEEDED })
}

export function installFakePowerPoint(model: FakePresentation, options: InstallOptions = {}): FakePowerPointHandle {
  if (options.apiVersion) model.apiVersion = options.apiVersion
  let runs = 0
  let syncs = 0
  const run = async <T>(batch: (context: unknown) => Promise<T>): Promise<T> => {
    runs += 1
    const ctx = new Ctx(() => model.apiVersion)
    const context = {
      presentation: new PresentationObj(ctx, model),
      sync: async () => {
        try {
          await ctx.sync()
        } finally {
          syncs += 1
        }
      },
      trackedObjects: { add: () => undefined, remove: () => undefined },
    }
    const result = await batch(context)
    await context.sync()
    return result
  }

  const enums = Object.fromEntries(
    ['GeometricShapeType', 'ShapeType', 'ConnectorType', 'ShapeZOrder', 'ShapeAutoSize', 'TextVerticalAlignment', 'ParagraphHorizontalAlignment',
      'PlaceholderType', 'InsertSlideFormatting', 'ShapeFillType', 'ShapeLineDashStyle', 'ShapeLineStyle', 'TableStyle', 'ShapeFontUnderlineStyle',
    ].map((name) => [name, pascalEnum()]),
  )

  const document = {
    get url() { return '' },
    setSelectedDataAsync: (data: string, optionsOrCallback?: unknown, callback?: Callback) => setSelectedDataAsync(model, data, optionsOrCallback, callback),
    addHandlerAsync: (eventType: string, handler: (args: unknown) => unknown, optionsOrCallback?: unknown, maybeCallback?: Callback) => {
      model.handlers.push({ eventType, handler })
      const callback = (typeof optionsOrCallback === 'function' ? optionsOrCallback : maybeCallback) as Callback | undefined
      callback?.({ status: SUCCEEDED })
    },
    removeHandlerAsync: (eventType: string, optionsOrCallback?: unknown, maybeCallback?: Callback) => {
      const options = (typeof optionsOrCallback === 'object' && optionsOrCallback !== null ? optionsOrCallback : {}) as { handler?: unknown }
      model.handlers = model.handlers.filter((h) => h.eventType !== eventType || (options.handler !== undefined && h.handler !== options.handler))
      const callback = (typeof optionsOrCallback === 'function' ? optionsOrCallback : maybeCallback) as Callback | undefined
      callback?.({ status: SUCCEEDED })
    },
    settings: {
      get: (key: string) => model.settings.get(key) ?? null,
      set: (key: string, value: unknown) => void model.settings.set(key, value),
      remove: (key: string) => void model.settings.delete(key),
      saveAsync: (callback?: Callback) => callback?.({ status: SUCCEEDED }),
    },
  }

  const g = globalThis as Record<string, unknown>
  const createPresentation = async (base64File?: string): Promise<void> => {
    model.createdPresentations.push(base64File ?? null)
  }
  g.PowerPoint = { run, createPresentation, ...enums }
  g.OfficeExtension = { Error: FakeOfficeError }
  g.Office = {
    HostType: { Word: 'Word', Excel: 'Excel', PowerPoint: 'PowerPoint', Outlook: 'Outlook', OneNote: 'OneNote', Project: 'Project', Access: 'Access' },
    CoercionType: { Text: 'text', Matrix: 'matrix', Table: 'table', Html: 'html', Ooxml: 'ooxml', SlideRange: 'slideRange', Image: 'image', XmlSvg: 'xmlSvg' },
    EventType: { DocumentSelectionChanged: 'documentSelectionChanged', ActiveViewChanged: 'activeViewChanged' },
    AsyncResultStatus: { Succeeded: SUCCEEDED, Failed: FAILED },
    context: {
      host: 'PowerPoint',
      requirements: {
        isSetSupported: (name: string, version = '1.1') => {
          const newest = name === 'PowerPointApi' ? model.apiVersion : model.otherSets[name]
          return newest !== undefined && compareVersions(newest, String(version)) >= 0
        },
      },
      document,
    },
    onReady: async () => ({ host: 'PowerPoint', platform: 'PC' }),
  }
  return { model, runs: () => runs, syncs: () => syncs }
}

export function uninstallFakePowerPoint(): void {
  const g = globalThis as Record<string, unknown>
  delete g.PowerPoint
  delete g.OfficeExtension
  delete g.Office
}

/** Fires DocumentSelectionChanged and waits for every handler, including any promise a handler returns. */
export async function fireSelectionChanged(model: FakePresentation): Promise<void> {
  const g = globalThis as { Office?: { context?: { document?: unknown } } }
  const handlers = model.handlers.filter((h) => h.eventType === 'documentSelectionChanged')
  await Promise.all(handlers.map(async (h) => h.handler({ type: 'documentSelectionChanged', document: g.Office?.context?.document })))
}

/** Changes the selection the way a click does: selects the shapes, then fires DocumentSelectionChanged. */
export async function clickShapes(model: FakePresentation, shapeIds: string[], slideId?: string): Promise<void> {
  const id = slideId ?? model.selection.slideIds[0] ?? model.slides[0]?.id
  model.selection = { slideIds: id === undefined ? [] : [id], shapeIds: [...shapeIds], textRange: null }
  await fireSelectionChanged(model)
}
