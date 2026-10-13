// What the slide tools and the deck check add to src/ppt/run.ts: the API version each feature needs, batched loads
// that survive one unreadable shape, and finding a shape by its path through groups.

/** The PowerPoint JavaScript API version each feature needs. */
export const NEEDS = {
  /** Slide tags. */
  tags: '1.3',
  /** Shapes with their text, position and formatting; new text boxes. */
  read: '1.4',
  /** The current slide, and going to a slide or shape. */
  current: '1.5',
  /** Moving a new slide next to the current one; placeholder types; shapes inside groups. */
  move: '1.8',
  /** Copying slides to a new presentation; alt text; the slide size. */
  export: '1.10',
} as const

/**
 * Loads what `queue` asks for, for many items in one sync. One item PowerPoint cannot read fails the whole batch, so
 * on failure each item is queued again in a sync of its own; the items that still fail are returned.
 */
export async function syncEach<T>(context: PowerPoint.RequestContext, items: T[], queue: (item: T) => void): Promise<Set<T>> {
  const failed = new Set<T>()
  if (!items.length) return failed
  try {
    items.forEach(queue)
    await context.sync()
    return failed
  } catch {
    // Fall through to one item at a time.
  }
  for (const item of items) {
    try {
      queue(item)
      await context.sync()
    } catch {
      failed.add(item)
    }
  }
  return failed
}

/** A shape by its path of ids from the slide's top level down through any groups. */
export function shapeAt(slide: PowerPoint.Slide, path: string[]): PowerPoint.Shape {
  const [first = '', ...rest] = path
  let shape = slide.shapes.getItem(first)
  for (const id of rest) shape = shape.group.shapes.getItem(id)
  return shape
}
