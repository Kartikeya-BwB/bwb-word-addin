// Fake shape adjustments (PowerPointApi 1.10): the handles that set, for example, how round a rounded rectangle's
// corners are. Office.js does not document their units; the fake keeps whatever numbers it is given.

import { Ctx, FakeClientResult, FakeObject, FakeOfficeError } from './base.ts'
import type { FakeShape } from './model.ts'

/** Adjustments a new shape of each type starts with; other types have none. */
const DEFAULTS: Record<string, number[]> = {
  RoundRectangle: [0.16667],
  Chevron: [0.5],
  HomePlate: [0.5],
  Trapezoid: [0.25],
  Donut: [0.25],
  RightArrow: [0.5, 0.5],
}

export function defaultAdjustments(geometricShapeType: string): number[] {
  return [...(DEFAULTS[geometricShapeType] ?? [])]
}

export class AdjustmentsObj extends FakeObject {
  constructor(context: Ctx, private readonly target: () => FakeShape) {
    super(context)
  }
  private values(): number[] {
    const s = this.target()
    s.adjustments ??= []
    return s.adjustments
  }
  get count(): number { this.need('count'); return this.values().length }
  get(index: number): FakeClientResult<number> {
    return new FakeClientResult(this.context, () => {
      const v = this.values()[index]
      if (v === undefined) throw new FakeOfficeError('InvalidArgument', `No adjustment ${index}`)
      return v
    })
  }
  set(index: number, value: number): void {
    this.write(() => {
      const list = this.values()
      if (index < 0 || index >= list.length) throw new FakeOfficeError('InvalidArgument', `No adjustment ${index}`)
      list[index] = value
    })
  }
}
