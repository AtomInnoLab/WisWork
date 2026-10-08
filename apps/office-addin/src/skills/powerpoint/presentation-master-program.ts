import type {
  PowerPointMasterOperation,
  PowerPointMasterState,
} from './browser-powerpoint-adapter.js'

export const MASTER_PATTERN_TYPES = [
  'Percent5',
  'Percent10',
  'Percent20',
  'Percent25',
  'Percent30',
  'Percent40',
  'Percent50',
  'Percent60',
  'Percent70',
  'Percent75',
  'Percent80',
  'Percent90',
  'Horizontal',
  'Vertical',
  'LightHorizontal',
  'LightVertical',
  'DarkHorizontal',
  'DarkVertical',
  'NarrowHorizontal',
  'NarrowVertical',
  'DashedHorizontal',
  'DashedVertical',
  'Cross',
  'DownwardDiagonal',
  'UpwardDiagonal',
  'LightDownwardDiagonal',
  'LightUpwardDiagonal',
  'DarkDownwardDiagonal',
  'DarkUpwardDiagonal',
  'WideDownwardDiagonal',
  'WideUpwardDiagonal',
  'DashedDownwardDiagonal',
  'DashedUpwardDiagonal',
  'DiagonalCross',
  'SmallCheckerBoard',
  'LargeCheckerBoard',
  'SmallGrid',
  'LargeGrid',
  'DottedGrid',
  'SmallConfetti',
  'LargeConfetti',
  'HorizontalBrick',
  'DiagonalBrick',
  'SolidDiamond',
  'OutlinedDiamond',
  'DottedDiamond',
  'Plaid',
  'Sphere',
  'Weave',
  'Divot',
  'Shingle',
  'Wave',
  'Trellis',
  'ZigZag',
] as const

export function projectedMasterState(
  before: PowerPointMasterState,
  operations: PowerPointMasterOperation[],
): PowerPointMasterState {
  const value = structuredClone(before)
  for (const operation of operations) {
    const master = value.masters.find((item) => item.id === operation.master_id)
    if (!master) throw new Error('invalid_tool_input')
    if (operation.op === 'set_master_background') {
      if (operation.fill.type === 'solid')
        master.background = {
          type: 'Solid',
          color: operation.fill.color,
          transparency: operation.fill.transparency,
        }
      else if (operation.fill.type === 'gradient')
        master.background = {
          type: 'Gradient',
          gradientType: operation.fill.gradient_type,
        } as PowerPointMasterState['masters'][number]['background']
      else if (operation.fill.type === 'pattern')
        master.background = {
          type: 'Pattern',
          pattern: operation.fill.pattern,
          foregroundColor: operation.fill.foreground_color,
          backgroundColor: operation.fill.background_color,
        } as PowerPointMasterState['masters'][number]['background']
      else
        master.background = {
          type: 'PictureOrTexture',
          pictureTransparency: operation.fill.transparency,
        }
    } else if (operation.op === 'set_master_theme_color')
      master.themeColors[operation.theme_color] = operation.color
    else {
      const layout = master.layouts.find((item) => item.id === operation.layout_id)
      if (!layout) throw new Error('invalid_tool_input')
      layout.isMasterBackgroundFollowed = operation.follow_master
      layout.areBackgroundGraphicsHidden = !operation.show_master_graphics
    }
  }
  return value
}

export function masterOperationKey(operation: PowerPointMasterOperation): string {
  if (operation.op === 'set_master_background') return `${operation.master_id}:background`
  if (operation.op === 'set_master_theme_color')
    return `${operation.master_id}:theme:${operation.theme_color}`
  return `${operation.master_id}:layout:${operation.layout_id}:background-following`
}

function normalizedColor(value: unknown): unknown {
  return typeof value === 'string' && /^#[0-9a-f]{6}$/i.test(value) ? value.toUpperCase() : value
}

export function masterOperationValue(
  state: PowerPointMasterState,
  operation: PowerPointMasterOperation,
): unknown {
  const master = state.masters.find((item) => item.id === operation.master_id)
  if (!master) return undefined
  if (operation.op === 'set_master_background') {
    const background = master.background
    return {
      type: background.type.toLowerCase(),
      ...(background.color === undefined ? {} : { color: normalizedColor(background.color) }),
      ...(background.transparency === undefined ? {} : { transparency: background.transparency }),
      ...(background.gradientType === undefined
        ? {}
        : { gradientType: background.gradientType.toLowerCase() }),
      ...(background.pattern === undefined ? {} : { pattern: background.pattern.toLowerCase() }),
      ...(background.foregroundColor === undefined
        ? {}
        : { foregroundColor: normalizedColor(background.foregroundColor) }),
      ...(background.backgroundColor === undefined
        ? {}
        : { backgroundColor: normalizedColor(background.backgroundColor) }),
      ...(background.pictureTransparency === undefined
        ? {}
        : { pictureTransparency: background.pictureTransparency }),
    }
  }
  if (operation.op === 'set_master_theme_color')
    return normalizedColor(master.themeColors[operation.theme_color])
  const layout = master.layouts.find((item) => item.id === operation.layout_id)
  return layout
    ? {
        follow_master: layout.isMasterBackgroundFollowed,
        show_master_graphics: !layout.areBackgroundGraphicsHidden,
      }
    : undefined
}

export function sameMasterOperationValue(
  actual: PowerPointMasterState,
  expected: PowerPointMasterState,
  operation: PowerPointMasterOperation,
): boolean {
  return (
    JSON.stringify(masterOperationValue(actual, operation)) ===
    JSON.stringify(masterOperationValue(expected, operation))
  )
}

export function inverseMasterOperation(
  before: PowerPointMasterState,
  operation: PowerPointMasterOperation,
): PowerPointMasterOperation {
  const master = before.masters.find((item) => item.id === operation.master_id)
  if (!master) throw new Error('invalid_tool_input')
  if (operation.op === 'set_master_background') {
    const type = master.background.type.toLowerCase()
    if (
      type === 'solid' &&
      typeof master.background.color === 'string' &&
      typeof master.background.transparency === 'number'
    )
      return {
        op: 'set_master_background',
        master_id: operation.master_id,
        fill: {
          type: 'solid',
          color: master.background.color,
          transparency: master.background.transparency,
        },
      }
    if (type === 'gradient' && typeof master.background.gradientType === 'string')
      return {
        op: 'set_master_background',
        master_id: operation.master_id,
        fill: { type: 'gradient', gradient_type: master.background.gradientType },
      }
    if (
      type === 'pattern' &&
      typeof master.background.pattern === 'string' &&
      typeof master.background.foregroundColor === 'string' &&
      typeof master.background.backgroundColor === 'string'
    )
      return {
        op: 'set_master_background',
        master_id: operation.master_id,
        fill: {
          type: 'pattern',
          pattern: master.background.pattern,
          foreground_color: master.background.foregroundColor,
          background_color: master.background.backgroundColor,
        },
      }
    throw new Error('office_api_unsupported')
  }
  if (operation.op === 'set_master_theme_color') {
    const color = master.themeColors[operation.theme_color]
    if (!color) throw new Error('office_api_unsupported')
    return {
      op: operation.op,
      master_id: operation.master_id,
      theme_color: operation.theme_color,
      color,
    }
  }
  const layout = master.layouts.find((item) => item.id === operation.layout_id)
  if (!layout) throw new Error('invalid_tool_input')
  return {
    op: operation.op,
    master_id: operation.master_id,
    layout_id: operation.layout_id,
    follow_master: layout.isMasterBackgroundFollowed,
    show_master_graphics: !layout.areBackgroundGraphicsHidden,
  }
}

/** Logical native fields; layout fill type is derived when inheritance changes. */
export function masterStateValuesFingerprint(state: PowerPointMasterState): string {
  const compare = (a: { id: string }, b: { id: string }) => (a.id < b.id ? -1 : a.id > b.id ? 1 : 0)
  return JSON.stringify(
    [...state.masters].sort(compare).map((master) => ({
      id: master.id,
      name: master.name,
      background: masterOperationValue(state, {
        op: 'set_master_background',
        master_id: master.id,
        fill: { type: 'solid', color: '#000000', transparency: 0 },
      }),
      themeColors: Object.fromEntries(
        Object.entries(master.themeColors)
          .sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0))
          .map(([key, value]) => [key, normalizedColor(value)]),
      ),
      layouts: [...master.layouts].sort(compare).map((layout) => ({
        id: layout.id,
        name: layout.name,
        follow_master: layout.isMasterBackgroundFollowed,
        show_master_graphics: !layout.areBackgroundGraphicsHidden,
      })),
    })),
  )
}
