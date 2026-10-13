// Types shared by the geometry engine. Everything is in points, as PowerPoint reports shapes.

/**
 * A shape's frame on the slide: the top-left corner and size of the unrotated frame, as PowerPoint reports them, and
 * the rotation in degrees clockwise about the frame's centre.
 */
export interface Box {
  id: string
  left: number
  top: number
  width: number
  height: number
  rotation?: number
}

export interface Rect {
  left: number
  top: number
  width: number
  height: number
}

export interface SlideSize {
  width: number
  height: number
}

export type Axis = 'horizontal' | 'vertical'

/**
 * What shapes line up with: the selection's outer edges, the anchor (the key
 * object, which never moves), the slide edges, the BwB margins or the content area between the title and the source.
 */
export type Reference = 'selection' | 'anchor' | 'slide' | 'margins' | 'content'

/** Where distribute spreads shapes: between the outer shapes of the selection, or across an area of the slide. */
export type Extent = 'selection' | 'slide' | 'margins' | 'content'

export type Edge = 'left' | 'right' | 'top' | 'bottom'

/** An operation that cannot run on what was given. The message is written for the person using the add-in. */
export class GeometryError extends Error {
  override name = 'GeometryError'
}
