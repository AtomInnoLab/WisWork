/**
 * Full-sheet Find coverage for streamed (lazy) workbooks.
 *
 * Univer's Find dialog searches the in-memory cell matrix, but a streamed
 * workbook only holds rows that were already scrolled into view, so matches
 * in never-visited rows are invisible and users conclude the data does not
 * exist (issue #113). This module wraps the built-in sheets find provider:
 * the inner model keeps handling everything inside the loaded window, while
 * the wrapper extends the session with out-of-window matches paged from the
 * underlying file via readSheetRangeMapped (journal edits included) — the
 * same approach the AI-side find takes. Focusing an out-of-window match
 * activates its sheet, starts loading its range, scrolls to it, and selects
 * it, so the grid shows real data instead of an empty jump.
 */
import type { ICellData, IRange, Workbook } from '@univerjs/core'
import {
  escapeRegExp,
  ICommandService,
  IUniverInstanceService,
  ObjectMatrix,
  replaceInDocumentBody,
  Tools,
} from '@univerjs/core'
import {
  FindBy,
  FindModel,
  IFindReplaceService,
  type IFindMatch,
  type IFindMoveParams,
  type IFindQuery,
  type IFindReplaceProvider,
  type IReplaceAllResult,
} from '@univerjs/find-replace'
import { SheetReplaceCommand } from '@univerjs/sheets-find-replace'
import { Subject, type Subscription } from 'rxjs'
export const LAZY_FIND_BATCH_CELLS = 18_000
export const LAZY_FIND_MAX_SCAN_CELLS = 400_000
export const LAZY_FIND_MAX_MATCHES = 10_000
export const LAZY_FIND_MAX_BATCHES = 64
export const LAZY_FIND_MAX_MS = 10_000
export const LAZY_FIND_MAX_BYTES = 16 * 1024 * 1024
export const LAZY_FIND_MAX_QUERY_LENGTH = 4_096
import { t } from './i18n/locale'
import { netAxisDelta } from './view-transform'
import type { LazyWorkbookState, UniverRuntime } from './univer-state'
import { ensureLazyRangeLoaded, readSheetRangeMapped } from './univer-sync'

/** Same match shape the built-in sheets provider produces (ISheetCellMatch). */
export interface LazyCellMatch extends IFindMatch {
  isFormula: boolean
  replaceable?: boolean
  /// Extra bookkeeping the wrapper needs to focus/replace the hit; ignored
  /// by Univer's composite model.
  range: { subUnitId: string; range: IRange }
  matchedText?: string | null
}

export interface LazyCellTexts {
  /** Display/computed value stringified like Univer's extractPureValue. */
  value: string | null
  formula: string | undefined
}

type LazyCellTest = (cell: LazyCellTexts) => boolean

/** Whether a find session needs the file-backed extension for this workbook. */
export function planLazyFind(state: LazyWorkbookState | null): 'inactive' | 'extend' {
  if (!state || state.flags.preloadComplete) return 'inactive'
  return 'extend'
}

/// Stringifies a scalar like Univer's extractPureValue: numbers become their
/// decimal text, booleans become "1"/"0".
export function scalarToText(value: unknown): string | null {
  if (value === null || value === undefined) return null
  if (typeof value === 'number') return `${value}`
  if (typeof value === 'boolean') return value ? '1' : '0'
  return String(value)
}

/// Mirrors the built-in provider's _preprocessQuery: lowercase unless
/// case-sensitive, then trim.
export function preprocessNeedle(query: IFindQuery): string {
  const raw = query.findString ?? ''
  return (query.caseSensitive === true ? raw : raw.toLowerCase()).trim()
}

/// Mirrors Univer's matchCellData/hitCell semantics for file-backed cells:
/// substring vs whole-cell (spaces trimmed, line breaks kept), case
/// sensitivity, and formula-vs-value look-in.
export function buildLazyCellTest(query: IFindQuery): LazyCellTest | null {
  if ((query.findString?.length ?? 0) > LAZY_FIND_MAX_QUERY_LENGTH) return null
  const caseSensitive = query.caseSensitive === true
  // The built-in provider preprocesses the query once up front (lowercase
  // unless case-sensitive, then trim); mirror it so both models agree on
  // what a hit is.
  const needle = preprocessNeedle(query)
  if (!needle) return null
  const matches = (text: string | null | undefined): boolean => {
    if (text === null || text === undefined) return false
    const haystack = caseSensitive ? text : text.toLowerCase()
    if (query.matchesTheWholeCell) {
      const trimmed = haystack.replace(/^ +/g, '').replace(/ +$/g, '')
      return trimmed === needle
    }
    return haystack.includes(needle)
  }
  return ({ value, formula }) => {
    if (formula && query.findBy === FindBy.FORMULA) return matches(formula)
    return matches(value)
  }
}

function insideRange(range: IRange | undefined, row: number, column: number): boolean {
  if (!range) return false
  return (
    row >= range.startRow &&
    row <= range.endRow &&
    column >= range.startColumn &&
    column <= range.endColumn
  )
}

function subtractRange(range: IRange, blocker: IRange): IRange[] {
  const top = Math.max(range.startRow, blocker.startRow)
  const bottom = Math.min(range.endRow, blocker.endRow)
  const left = Math.max(range.startColumn, blocker.startColumn)
  const right = Math.min(range.endColumn, blocker.endColumn)
  if (top > bottom || left > right) return [range]
  const pieces: IRange[] = []
  if (range.startRow < top) pieces.push({ ...range, endRow: top - 1 })
  if (bottom < range.endRow) pieces.push({ ...range, startRow: bottom + 1 })
  if (range.startColumn < left) {
    pieces.push({
      startRow: top,
      endRow: bottom,
      startColumn: range.startColumn,
      endColumn: left - 1,
    })
  }
  if (right < range.endColumn) {
    pieces.push({
      startRow: top,
      endRow: bottom,
      startColumn: right + 1,
      endColumn: range.endColumn,
    })
  }
  return pieces
}

function disjointRanges(ranges: readonly IRange[]): IRange[] {
  const accepted: IRange[] = []
  for (const range of ranges) {
    let pieces = [{ ...range }]
    for (const blocker of accepted)
      pieces = pieces.flatMap((piece) => subtractRange(piece, blocker))
    accepted.push(...pieces)
  }
  return accepted
}

function rangesEqual(left: IRange, right: IRange): boolean {
  return (
    left.startRow === right.startRow &&
    left.endRow === right.endRow &&
    left.startColumn === right.startColumn &&
    left.endColumn === right.endColumn
  )
}

/** True when the loaded window covers the coordinate — the inner model owns it. */
export function coveredByWindow(
  state: LazyWorkbookState,
  sheetId: string,
  row: number,
  column: number,
): boolean {
  return insideRange(state.loadedRanges.get(sheetId), row, column)
}

function matchKey(match: LazyCellMatch): string {
  return `${match.range.subUnitId}|${match.range.range.startRow}|${match.range.range.startColumn}`
}

/** Dedupes inner-model matches against the extension's; inner entries win. */
export function mergeFindMatches(primary: IFindMatch[], extra: LazyCellMatch[]): IFindMatch[] {
  if (extra.length === 0) return primary
  const seen = new Set(primary.map((match) => matchKey(match as LazyCellMatch)))
  const merged = [...primary]
  for (const match of extra) {
    const key = matchKey(match)
    if (seen.has(key)) continue
    seen.add(key)
    merged.push(match)
  }
  return merged
}

interface ScanCell {
  readonly row: number
  readonly column: number
  readonly value: string | number | boolean | null
  readonly formula: string | undefined
}

/** Journal edits of one sheet that sit outside the loaded window. */
export function collectJournalMatches(
  state: LazyWorkbookState,
  sheetId: string,
  test: LazyCellTest,
  limit = Number.POSITIVE_INFINITY,
  accept: (row: number, column: number) => boolean = () => true,
): ScanCell[] {
  const found: ScanCell[] = []
  if (limit <= 0) return found
  const journal = state.editJournal.cells.get(sheetId)
  for (const entry of journal?.values() ?? []) {
    if (!entry.hasValue) continue
    if (!accept(entry.row, entry.column)) continue
    if (coveredByWindow(state, sheetId, entry.row, entry.column)) continue
    if (!test({ value: scalarToText(entry.value), formula: entry.formula })) continue
    found.push({ row: entry.row, column: entry.column, value: entry.value, formula: entry.formula })
    if (found.length >= limit) break
  }
  return found
}

/** Coordinates whose file cell is shadowed by a journal edit this session.
 *  Style-only entries (hasValue false) leave the file's content authoritative
 *  and must stay findable. */
export function journalShadowKeys(state: LazyWorkbookState, sheetId: string): Set<string> {
  const shadowed = new Set<string>()
  const journal = state.editJournal.cells.get(sheetId)
  for (const entry of journal?.values() ?? []) {
    if (entry.hasValue) shadowed.add(`${entry.row}:${entry.column}`)
  }
  return shadowed
}

/** Row-major/column-major ordering across sheets, following the query direction. */
export function extraComparator(
  sheetOrder: ReadonlyMap<string, number>,
  columnDirection: boolean,
): (a: ScanCell & { sheetId: string }, b: ScanCell & { sheetId: string }) => number {
  return (a, b) => {
    const sheetDelta = (sheetOrder.get(a.sheetId) ?? 0) - (sheetOrder.get(b.sheetId) ?? 0)
    if (sheetDelta !== 0) return sheetDelta
    return columnDirection
      ? a.column - b.column || a.row - b.row
      : a.row - b.row || a.column - b.column
  }
}

function makeCellMatch(
  unitId: string,
  sheetId: string,
  cell: ScanCell,
  findByFormula: boolean,
): LazyCellMatch {
  const isFormula = Boolean(cell.formula)
  return {
    provider: 'sheets-find-replace-provider',
    unitId,
    isFormula,
    // Formula hits are only replaced when searching formulas (mirrors the
    // built-in model); plain cells behave exactly like in-memory ones.
    replaceable: isFormula ? findByFormula : cell.value !== null && cell.value !== undefined,
    matchedText: (findByFormula && isFormula ? cell.formula : scalarToText(cell.value)) ?? null,
    range: {
      subUnitId: sheetId,
      range: {
        startRow: cell.row,
        endRow: cell.row,
        startColumn: cell.column,
        endColumn: cell.column,
      },
    },
  }
}

export interface LazyFindBridgeDeps {
  runtime: UniverRuntime
  lazyWorkbookRef: { current: LazyWorkbookState | null }
  setMessage: (message: string) => void
}

interface InnerFindModel extends FindModel {
  readonly unitId: string
  focusSelection(): void
}

/**
 * Replaces the registered sheets find provider with a wrapper while the app
 * lives. Non-streamed workbooks flow straight through; streamed ones get the
 * extended model. On dispose the adopted providers go back into the service.
 */
export function installLazyFindBridge(deps: LazyFindBridgeDeps): { dispose(): void } {
  const service = deps.runtime.univer.__getInjector().get(IFindReplaceService)
  const providers = service.getProviders()
  const adopted = new Set<IFindReplaceProvider>()
  let generation = 0
  // Registering only APPENDS to the service's live provider set, and
  // _startSearching dispatches to every provider in it — leaving the built-in
  // registered would run it twice per search (double-counted in-window
  // matches, and its second find() disposes the first session's model).
  // Detach built-ins into `adopted` so the service reaches them only through
  // the wrapper; re-sweep on every call in case one registered later.
  const adoptForeign = () => {
    for (const provider of [...providers]) {
      if (provider === wrapper) continue
      providers.delete(provider)
      adopted.add(provider)
    }
  }
  const wrapper: IFindReplaceProvider = {
    async find(query: IFindQuery) {
      generation += 1
      const liveGeneration = generation
      adoptForeign()
      const models: FindModel[] = []
      for (const builtin of [...adopted]) {
        models.push(...(await builtin.find(query)))
      }
      const state = deps.lazyWorkbookRef.current
      if (!state || planLazyFind(state) !== 'extend' || models.length === 0) return models
      return models.map(
        (model) =>
          new LazyExtendedFindModel(
            model as InnerFindModel,
            state,
            query,
            deps,
            () => generation === liveGeneration,
          ),
      )
    },
    terminate() {
      generation += 1
      adoptForeign()
      for (const builtin of adopted) builtin.terminate()
    },
  }
  const registration = service.registerFindReplaceProvider(wrapper)
  adoptForeign()
  return {
    dispose() {
      generation += 1
      registration.dispose()
      for (const builtin of adopted) providers.add(builtin)
      adopted.clear()
    },
  }
}

/**
 * A FindModel combining the built-in in-window session with out-of-window
 * hits paged from the underlying file. The inner model keeps navigating and
 * highlighting everything it can see; this wrapper only steps in when the
 * inner session runs out, and hands focus back once the jumped-to region is
 * materialized in the grid (the inner model re-runs on mutations and takes
 * over navigation there).
 */
export class LazyExtendedFindModel extends FindModel {
  readonly unitId: string

  readonly matchesUpdate$ = new Subject<IFindMatch[]>()
  readonly activelyChangingMatch$ = new Subject<LazyCellMatch>()

  private readonly state: LazyWorkbookState
  private readonly query: IFindQuery
  private readonly deps: LazyFindBridgeDeps
  private readonly isLiveGeneration: () => boolean

  private alive = true
  private truncated = false
  private readonly scanAbort = new AbortController()
  private deadlineExpired = false
  private readonly deadlineTimer: ReturnType<typeof setTimeout>
  private selectionSignature: string
  private readonly selectionRanges: readonly IRange[] | null
  private extras: LazyCellMatch[] = []
  private lastFocusedExtra: LazyCellMatch | null = null
  private pendingInner: IFindMatch | null = null
  private readonly forwardSub: Subscription

  constructor(
    private readonly inner: InnerFindModel,
    state: LazyWorkbookState,
    query: IFindQuery,
    deps: LazyFindBridgeDeps,
    isLiveGeneration: () => boolean,
  ) {
    super()
    this.state = state
    this.query = query
    this.deps = deps
    this.isLiveGeneration = isLiveGeneration
    this.unitId = inner.unitId
    this.selectionSignature = this.currentSelectionSignature()
    this.selectionRanges = this.captureSelectionRanges()
    this.deadlineTimer = setTimeout(() => {
      if (!this.alive) return
      this.deadlineExpired = true
      this.scanAbort.abort()
    }, LAZY_FIND_MAX_MS)
    // The inner model refreshes itself when grid mutations stream regions in
    // or evict them; forward those moments so the dialog count stays right.
    this.forwardSub = inner.matchesUpdate$.subscribe(() => this.emitMerged())
    void this.runScan()
  }

  override dispose(): void {
    this.alive = false
    clearTimeout(this.deadlineTimer)
    this.scanAbort.abort()
    this.forwardSub.unsubscribe()
    this.matchesUpdate$.complete()
    this.activelyChangingMatch$.complete()
    super.dispose()
  }

  getMatches(): IFindMatch[] {
    if (!this.stateIsCurrent()) return []
    return mergeFindMatches(this.innerMatches(), this.currentExtras())
  }

  moveToNextMatch(params?: IFindMoveParams): LazyCellMatch | null {
    if (!this.stateIsCurrent()) return null
    return this.moveThroughMatches('next', params)
  }

  moveToPreviousMatch(params?: IFindMoveParams): LazyCellMatch | null {
    if (!this.stateIsCurrent()) return null
    return this.moveThroughMatches('previous', params)
  }

  /**
   * Segmented cursor: the inner (index-cursor) session first, then the
   * extras, then wrap back to the inner. The upstream `loop` must never
   * reach the inner model — its loop takes the modulo of its own match list
   * and would cycle in-window hits forever, starving the extras. When the
   * inner runs out it resets its own cursor, so the wrap-around re-entry
   * uses ignoreSelection to land on its first/last match. Boundary order is
   * not strict document order (a mid-sheet window hands over to the
   * top-most extra) — accepted simplification.
   */
  private moveThroughMatches(
    direction: 'next' | 'previous',
    params?: IFindMoveParams,
  ): LazyCellMatch | null {
    if (this.pendingInner) {
      const pendingKey = matchKey(this.pendingInner as LazyCellMatch)
      this.pendingInner =
        this.innerMatches().find(
          (candidate) => matchKey(candidate as LazyCellMatch) === pendingKey,
        ) ?? null
    }
    if (this.lastFocusedExtra && this.pendingInner) {
      const nextExtra = this.neighborExtra(direction)
      if (nextExtra && this.extraPrecedes(nextExtra, this.pendingInner, direction)) {
        if (!params?.noFocus) this.focusExtra(nextExtra)
        else this.lastFocusedExtra = nextExtra
        return nextExtra
      }
      const pending = this.pendingInner
      this.pendingInner = null
      this.lastFocusedExtra = null
      if (!params?.noFocus) this.safeInnerFocus()
      return pending as LazyCellMatch
    }
    if (!this.lastFocusedExtra) {
      const candidate = this.innerNeighbor(direction, params)
      const extra = this.neighborExtra(direction)
      if (extra && (!candidate || this.extraPrecedes(extra, candidate, direction))) {
        this.pendingInner = candidate
        if (!params?.noFocus) this.focusExtra(extra)
        else this.lastFocusedExtra = extra
        return extra
      }
      if (candidate) {
        if (!params?.noFocus) this.safeInnerFocus()
        return candidate as LazyCellMatch
      }
    }
    const target = this.neighborExtra(direction)
    if (target) {
      if (!params?.noFocus) this.focusExtra(target)
      else this.lastFocusedExtra = target
      return target
    }
    if (params?.loop === false) return null
    this.lastFocusedExtra = null
    const wrapped = this.innerNeighbor(direction, { ...params, ignoreSelection: true })
    if (wrapped) {
      if (!params?.noFocus) this.safeInnerFocus()
      return wrapped as LazyCellMatch
    }
    // No inner matches at all — cycle within the extras themselves.
    const candidates = this.currentExtras()
    const first =
      direction === 'next' ? (candidates[0] ?? null) : (candidates[candidates.length - 1] ?? null)
    if (!first) return null
    if (!params?.noFocus) this.focusExtra(first)
    else this.lastFocusedExtra = first
    return first
  }

  private extraPrecedes(
    extra: LazyCellMatch,
    inner: IFindMatch,
    direction: 'next' | 'previous',
  ): boolean {
    const innerRange = (inner as LazyCellMatch).range
    if (!innerRange) return false
    const order = this.sheetOrderIndex()
    const position = (match: LazyCellMatch): [number, number, number] => {
      const bounds = match.range.range
      const primary = this.query.findDirection === 'column' ? bounds.startColumn : bounds.startRow
      const secondary = this.query.findDirection === 'column' ? bounds.startRow : bounds.startColumn
      return [order.get(match.range.subUnitId) ?? Number.MAX_SAFE_INTEGER, primary, secondary]
    }
    const delta = compareTriples(position(extra), position(inner as LazyCellMatch))
    return direction === 'next' ? delta < 0 : delta > 0
  }

  async replace(replaceString: string): Promise<boolean> {
    if (!this.stateIsCurrent()) return false
    // Only an extra that currently holds the segmented cursor may be
    // written; otherwise the inner session owns the current match.
    const extra = this.lastFocusedExtra
    if (extra) {
      if (extra.replaceable !== true) return false
      return this.writeExtraReplacement(extra, replaceString)
    }
    try {
      return await this.inner.replace(replaceString)
    } catch {
      return false
    }
  }

  async replaceAll(replaceString: string): Promise<IReplaceAllResult> {
    if (!this.stateIsCurrent()) return { success: 0, failure: 0 }
    if (this.selectionRanges) {
      return this.replaceAllInFrozenSelection(replaceString)
    }
    let success = 0
    let failure = 0
    try {
      const result = await this.inner.replaceAll(replaceString)
      success += result.success
      failure += result.failure
    } catch {
      /* the inner session may already be gone; still report the extension's */
    }
    for (const extra of this.currentExtras()) {
      if (extra.replaceable !== true) {
        failure += 1
        continue
      }
      if (await this.writeExtraReplacement(extra, replaceString)) success += 1
      else failure += 1
    }
    return { success, failure }
  }

  private async replaceAllInFrozenSelection(replaceString: string): Promise<IReplaceAllResult> {
    const seen = new Set<string>()
    const loaded = this.innerMatches().filter((match): match is LazyCellMatch => {
      const key = matchKey(match as LazyCellMatch)
      if (seen.has(key)) return false
      seen.add(key)
      return true
    })
    const extras = this.currentExtras().filter((match) => {
      const key = matchKey(match)
      if (seen.has(key)) return false
      seen.add(key)
      return true
    })

    let success = 0
    let failure = loaded.filter((match) => match.replaceable === false).length
    const replaceableLoaded = loaded.filter((match) => match.replaceable !== false)
    let native: IReplaceAllResult
    try {
      native = await this.replaceLoadedMatches(replaceableLoaded, replaceString)
    } catch {
      native = { success: 0, failure: replaceableLoaded.length }
    }
    success += native.success
    failure += native.failure

    for (const extra of extras) {
      if (extra.replaceable !== true) {
        failure += 1
      } else if (await this.writeExtraReplacement(extra, replaceString)) {
        success += 1
      } else {
        failure += 1
      }
    }
    return { success, failure }
  }

  private async replaceLoadedMatches(
    matches: LazyCellMatch[],
    replaceString: string,
  ): Promise<IReplaceAllResult> {
    if (matches.length === 0) return { success: 0, failure: 0 }
    const workbook = this.deps.runtime.univer
      .__getInjector()
      .get(IUniverInstanceService)
      .getUnit<Workbook>(this.unitId)
    if (!workbook) return { success: 0, failure: matches.length }

    const bySheet = new Map<string, LazyCellMatch[]>()
    for (const match of matches) {
      const sheetMatches = bySheet.get(match.range.subUnitId) ?? []
      sheetMatches.push(match)
      bySheet.set(match.range.subUnitId, sheetMatches)
    }
    const replacements: Array<{
      count: number
      subUnitId: string
      value: ReturnType<ObjectMatrix<ICellData>['getMatrix']>
    }> = []
    let failureCount = 0
    for (const [subUnitId, sheetMatches] of bySheet) {
      const worksheet = workbook.getSheetBySheetId(subUnitId)
      if (!worksheet) {
        failureCount += sheetMatches.length
        continue
      }
      const matrix = new ObjectMatrix<ICellData>()
      let count = 0
      for (const match of sheetMatches) {
        const { startRow, startColumn } = match.range.range
        const current = worksheet.getCellRaw(startRow, startColumn)
        const replacement = this.replacedCellData(match, current ?? null, replaceString)
        if (!replacement) {
          failureCount += 1
          continue
        }
        matrix.setValue(startRow, startColumn, replacement)
        count += 1
      }
      if (count > 0) replacements.push({ count, subUnitId, value: matrix.getMatrix() })
    }
    if (replacements.length === 0) return { success: 0, failure: failureCount }
    try {
      const result = (await this.deps.runtime.univer
        .__getInjector()
        .get(ICommandService)
        .executeCommand(SheetReplaceCommand.id, {
          unitId: this.unitId,
          replacements,
        })) as unknown as IReplaceAllResult
      return { success: result.success, failure: result.failure + failureCount }
    } catch {
      return { success: 0, failure: matches.length }
    }
  }

  private replacedCellData(
    match: LazyCellMatch,
    current: ICellData | null | undefined,
    replaceString: string,
  ): ICellData | null {
    if (!current) return null
    const needle = preprocessNeedle(this.query)
    if (!needle) return null
    const flags = this.query.caseSensitive ? 'g' : 'ig'
    if (match.isFormula) {
      if (this.query.findBy !== FindBy.FORMULA || !current.f) return null
      return {
        f: current.f.replace(new RegExp(escapeRegExp(needle), flags), replaceString),
        v: null,
      }
    }
    if (current.p?.body) {
      const richText = Tools.deepClone(current.p)
      const body = richText.body
      if (!body) return null
      replaceInDocumentBody(body, needle, replaceString, this.query.caseSensitive === true)
      return { p: richText }
    }
    if (current.v === null || current.v === undefined) return null
    return {
      v: current.v.toString().replace(new RegExp(escapeRegExp(needle), flags), replaceString),
    }
  }

  focusSelection(): void {
    if (this.lastFocusedExtra) {
      this.focusExtra(this.lastFocusedExtra)
      return
    }
    this.safeInnerFocus()
  }

  private safeInnerFocus(): void {
    try {
      this.inner.focusSelection()
    } catch {
      /* closed workbook */
    }
  }

  private innerNeighbor(
    direction: 'next' | 'previous',
    params?: IFindMoveParams,
  ): IFindMatch | null {
    // loop stays stripped: the inner model's own loop cycles its list
    // forever and would never yield to the extras.
    const stripped = { ...params, noFocus: true, loop: false }
    try {
      const maxSteps = Math.max(1, this.inner.getMatches().length + 1)
      for (let step = 0; step < maxSteps; step += 1) {
        const candidate =
          direction === 'next'
            ? this.inner.moveToNextMatch(stripped)
            : this.inner.moveToPreviousMatch(stripped)
        if (!candidate) return null
        if (this.matchInSelection(candidate)) return candidate
      }
      return null
    } catch {
      return null
    }
  }

  private innerMatches(): IFindMatch[] {
    try {
      const matches = this.inner.getMatches()
      return matches.filter((match) => this.matchInSelection(match))
    } catch {
      return []
    }
  }

  private matchInSelection(match: IFindMatch): boolean {
    if (!this.selectionRanges) return true
    const value = match as LazyCellMatch
    return this.selectionRanges.some((range) =>
      insideRange(range, value.range.range.startRow, value.range.range.startColumn),
    )
  }

  /** Extras that are still outside the (evolving) loaded window. */
  private currentExtras(): LazyCellMatch[] {
    return this.extras.filter((match) => {
      if (!this.sheetAllowed(match)) return false
      if (
        coveredByWindow(
          this.state,
          match.range.subUnitId,
          match.range.range.startRow,
          match.range.range.startColumn,
        )
      ) {
        return false
      }
      // The in-memory scan skips rows hidden by an active filter; hold
      // out-of-window hits to the same visibility rules. Fails open when the
      // filter state is not reachable (sheet gone mid-session).
      return !this.isRowHidden(match.range.subUnitId, match.range.range.startRow)
    })
  }

  /// Mirrors the built-in scan's worksheet.getRowFiltered check for cells
  /// that never streamed into Univer's matrix: the filter model lives on the
  /// workbook instance, not the grid, so it answers regardless of loading.
  private isRowHidden(subUnitId: string, row: number): boolean {
    try {
      const workbook = this.deps.runtime.univer
        .__getInjector()
        .get(IUniverInstanceService)
        .getUnit<Workbook>(this.unitId)
      return workbook?.getSheetBySheetId(subUnitId)?.getRowFiltered(row) === true
    } catch {
      return false
    }
  }

  /**
   * The first out-of-window hit after (or before) the current selection.
   * Exhaustion returns null — wrapping across the segments is
   * moveThroughMatches' job.
   */
  private neighborExtra(direction: 'next' | 'previous'): LazyCellMatch | null {
    const candidates = this.currentExtras()
    if (candidates.length === 0) return null
    const order = this.sheetOrderIndex()
    const columnDirection = this.query.findDirection === 'column'
    const axes = (row: number, column: number): [number, number] =>
      columnDirection ? [column, row] : [row, column]
    const sign = direction === 'next' ? 1 : -1
    const positionOf = (match: LazyCellMatch): [number, number, number] => {
      const bounds = match.range.range
      const [primary, secondary] = axes(bounds.startRow, bounds.startColumn)
      return [
        sign * (order.get(match.range.subUnitId) ?? Number.MAX_SAFE_INTEGER),
        sign * primary,
        sign * secondary,
      ]
    }
    const reference = this.referencePosition(order)
    if (!reference) {
      return direction === 'next' ? candidates[0]! : candidates[candidates.length - 1]!
    }
    const [referencePrimary, referenceSecondary] = axes(reference.row, reference.column)
    const referencePosition_: [number, number, number] = [
      sign * reference.sheetIndex,
      sign * referencePrimary,
      sign * referenceSecondary,
    ]
    const ordered = [...candidates].sort((a, b) => compareTriples(positionOf(a), positionOf(b)))
    return (
      ordered.find((match) => compareTriples(positionOf(match), referencePosition_) > 0) ?? null
    )
  }

  private sheetOrderIndex(): Map<string, number> {
    try {
      const sheets = this.deps.runtime.univerAPI.getActiveWorkbook()?.getSheets() ?? []
      return new Map(sheets.map((sheet, index) => [sheet.getSheetId(), index] as const))
    } catch {
      return new Map()
    }
  }

  private referencePosition(
    order: ReadonlyMap<string, number>,
  ): { sheetIndex: number; row: number; column: number } | null {
    if (this.lastFocusedExtra) {
      return {
        sheetIndex: order.get(this.lastFocusedExtra.range.subUnitId) ?? Number.MAX_SAFE_INTEGER,
        row: this.lastFocusedExtra.range.range.startRow,
        column: this.lastFocusedExtra.range.range.startColumn,
      }
    }
    try {
      const workbook = this.deps.runtime.univerAPI.getActiveWorkbook()
      const range = workbook?.getActiveRange()
      const activeSheet = workbook?.getActiveSheet()
      if (!workbook || !range || !activeSheet) return null
      return {
        sheetIndex: order.get(activeSheet.getSheetId()) ?? Number.MAX_SAFE_INTEGER,
        row: range.getRow(),
        column: range.getColumn(),
      }
    } catch {
      return null
    }
  }

  /** Guards against a session outliving its workbook: after a workbook
   *  switch, sheetIds may collide and getActiveWorkbook() targets the wrong
   *  book. */
  private stateIsCurrent(): boolean {
    if (this.deps.lazyWorkbookRef.current !== this.state) return false
    try {
      const active = this.deps.runtime.univerAPI.getActiveWorkbook()
      return (
        active?.getId() === this.unitId &&
        active.getId() === this.state.expectedWorkbookId &&
        this.currentSelectionSignature() === this.selectionSignature
      )
    } catch {
      return false
    }
  }

  private currentSelectionSignature(): string {
    try {
      const workbook = this.deps.runtime.univerAPI.getActiveWorkbook()
      if (!workbook) return 'unavailable'
      if (this.query.findScope === 'unit') {
        return JSON.stringify(
          workbook.getSheets().map((sheet) => [sheet.getSheetId(), sheet.isSheetHidden() === true]),
        )
      }
      const activeSheet = workbook.getActiveSheet()
      if (!activeSheet) return 'unavailable'
      const ranges = activeSheet.getSelection()?.getActiveRangeList() ?? []
      return JSON.stringify([
        activeSheet.getSheetId(),
        activeSheet.isSheetHidden() === true,
        ranges.map((range) => {
          const value = range.getRange()
          return [value.startRow, value.endRow, value.startColumn, value.endColumn]
        }),
      ])
    } catch {
      return 'unavailable'
    }
  }

  private captureSelectionRanges(): readonly IRange[] | null {
    if (this.query.findScope === 'unit') return null
    try {
      const worksheet = this.deps.runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
      const ranges =
        worksheet
          ?.getSelection()
          ?.getActiveRangeList()
          .map((range) => range.getRange()) ?? []
      const isSingle = (range: IRange): boolean => {
        const merged = worksheet?.getCellMergeData(range.startRow, range.startColumn)?.getRange()
        return merged
          ? rangesEqual(range, merged)
          : range.startRow === range.endRow && range.startColumn === range.endColumn
      }
      return ranges.some((range) => !isSingle(range)) ? disjointRanges(ranges) : null
    } catch {
      return null
    }
  }

  private sheetAllowed(match: LazyCellMatch): boolean {
    if (this.query.findScope !== 'unit') return true
    try {
      return (
        this.deps.runtime.univerAPI
          .getActiveWorkbook()
          ?.getSheetBySheetId(match.range.subUnitId)
          ?.isSheetHidden() !== true
      )
    } catch {
      return false
    }
  }

  private scanTargetIsCurrent(worksheet: { getSheetId(): string }): boolean {
    if (!this.stateIsCurrent()) return false
    try {
      return (
        this.deps.runtime.univerAPI
          .getActiveWorkbook()
          ?.getSheetBySheetId(worksheet.getSheetId()) === worksheet
      )
    } catch {
      return false
    }
  }

  /** Activate the sheet, load the region, scroll to it, and select the cell. */
  private focusExtra(match: LazyCellMatch): void {
    if (!this.stateIsCurrent()) return
    this.lastFocusedExtra = match
    try {
      const workbook = this.deps.runtime.univerAPI.getActiveWorkbook()
      if (!workbook) return
      const worksheet = workbook.getSheetBySheetId(match.range.subUnitId)
      if (!worksheet) return
      if (worksheet.getSheetId() !== workbook.getActiveSheet()?.getSheetId()) {
        workbook.setActiveSheet(worksheet)
      }
      const bounds = match.range.range
      // Best-effort streaming: the scroll below also triggers the regular
      // viewport load; this makes sure the exact hit lands even when the
      // visible window math picks a different anchor.
      void ensureLazyRangeLoaded(
        this.deps.runtime,
        this.deps.lazyWorkbookRef,
        worksheet,
        {
          startRow: bounds.startRow,
          endRow: bounds.endRow,
          startColumn: bounds.startColumn,
          endColumn: bounds.endColumn,
        },
        this.deps.setMessage,
      )
      worksheet.scrollToCell(bounds.startRow, bounds.startColumn)
      const target = worksheet.getRange(bounds.startRow, bounds.startColumn, 1, 1)
      if (this.selectionRanges) worksheet.getSelection()?.updatePrimaryCell(target)
      else target.activate()
      // Selection changes performed by Find itself are not external drift;
      // retain the original search ranges while advancing the live guard.
      this.selectionSignature = this.currentSelectionSignature()
      this.emitMerged()
      this.activelyChangingMatch$.next(match)
    } catch {
      /* closed workbook mid-jump */
    }
  }

  /** Writes the replacement straight onto the cell; the journal carries it. */
  private async writeExtraReplacement(
    match: LazyCellMatch,
    replaceString: string,
  ): Promise<boolean> {
    if (!this.stateIsCurrent()) return false
    try {
      const workbook = this.deps.runtime.univerAPI.getActiveWorkbook()
      const worksheet = workbook?.getSheetBySheetId(match.range.subUnitId)
      if (!worksheet) return false
      const bounds = match.range.range
      const target = worksheet.getRange(bounds.startRow, bounds.startColumn, 1, 1)
      const matchedText =
        match.matchedText ??
        (match.isFormula ? target.getFormula() : scalarToText(target.getValue())) ??
        ''
      if (match.isFormula) {
        target.setValues([[{ f: replaceAllOccurrences(matchedText, this.query, replaceString) }]])
      } else {
        target.setValues([[{ v: replaceAllOccurrences(matchedText, this.query, replaceString) }]])
      }
      return true
    } catch {
      return false
    }
  }

  private emitMerged(): void {
    if (!this.alive) return
    this.matchesUpdate$.next(this.getMatches())
  }

  /** Pages the underlying file for out-of-window hits, emitting as it goes. */
  private async runScan(): Promise<void> {
    const workbook = this.deps.runtime.univerAPI.getActiveWorkbook()
    if (!workbook || !this.stateIsCurrent()) return
    const test = buildLazyCellTest(this.query)
    if (!test) return
    const unitId = this.unitId
    const sheets = workbook.getSheets()
    const targets =
      this.query.findScope === 'unit'
        ? sheets.filter((sheet) => sheet.isSheetHidden() !== true)
        : sheets.filter((sheet) => sheet.getSheetId() === workbook.getActiveSheet()?.getSheetId())
    const sheetOrder = new Map(sheets.map((sheet, index) => [sheet.getSheetId(), index] as const))
    const comparator = extraComparator(sheetOrder, this.query.findDirection === 'column')
    const collected: (ScanCell & { sheetId: string })[] = []
    const collectedKeys = new Set<string>()
    // Budget counts scanned extent, not hits — the AI-side findInLazyWorkbook
    // semantics. Counting hits would scan sparse multi-million-cell sheets
    // end to end.
    let scannedCells = 0
    let scannedBatches = 0
    let scannedBytes = 0
    const startedAt = Date.now()

    for (const worksheet of targets) {
      const sheetId = worksheet.getSheetId()
      // Session edits first — they shadow file cells at the same coordinates.
      const journalMatches = collectJournalMatches(
        this.state,
        sheetId,
        test,
        Math.max(0, LAZY_FIND_MAX_MATCHES - collected.length),
        (row, column) =>
          !this.selectionRanges ||
          this.selectionRanges.some((range) => insideRange(range, row, column)),
      )
      for (const cell of journalMatches) {
        const key = `${sheetId}|${cell.row}|${cell.column}`
        if (collectedKeys.has(key)) continue
        collectedKeys.add(key)
        collected.push({ ...cell, sheetId })
      }
      if (collected.length >= LAZY_FIND_MAX_MATCHES) this.truncated = true
      const meta = this.state.file.sheets.find((candidate) => candidate.id === sheetId)
      // Sheets added this session live entirely in the journal.
      if (!meta || meta.rowCount <= 0 || meta.columnCount <= 0) {
        this.refreshExtras(collected, comparator, unitId)
        this.emitMerged()
        continue
      }
      const ops = this.state.editJournal.structuralOps.get(sheetId) ?? []
      const screenRows = Math.max(meta.rowCount + netAxisDelta(ops, 'row'), 0)
      const screenColumns = Math.max(meta.columnCount + netAxisDelta(ops, 'column'), 0)
      if (screenRows <= 0 || screenColumns <= 0) continue
      const scanRanges = (
        this.selectionRanges ?? [
          {
            startRow: 0,
            endRow: screenRows - 1,
            startColumn: 0,
            endColumn: screenColumns - 1,
          },
        ]
      )
        .map((range) => ({
          startRow: Math.max(0, range.startRow),
          endRow: Math.min(screenRows - 1, range.endRow),
          startColumn: Math.max(0, range.startColumn),
          endColumn: Math.min(screenColumns - 1, range.endColumn),
        }))
        .filter((range) => range.startRow <= range.endRow && range.startColumn <= range.endColumn)
      const shadowed = journalShadowKeys(this.state, sheetId)
      for (const scanRange of scanRanges) {
        const scanColumns = scanRange.endColumn - scanRange.startColumn + 1
        if (scanColumns > LAZY_FIND_BATCH_CELLS) {
          this.truncated = true
          break
        }
        const batchRows = Math.max(1, Math.floor(LAZY_FIND_BATCH_CELLS / scanColumns))
        for (
          let startRow = scanRange.startRow;
          startRow <= scanRange.endRow;
          startRow += batchRows
        ) {
          if (!this.alive || !this.isLiveGeneration() || !this.scanTargetIsCurrent(worksheet))
            return
          if (this.truncated) break
          if (
            scannedCells >= LAZY_FIND_MAX_SCAN_CELLS ||
            scannedBatches >= LAZY_FIND_MAX_BATCHES ||
            Date.now() - startedAt >= LAZY_FIND_MAX_MS ||
            collected.length >= LAZY_FIND_MAX_MATCHES
          ) {
            this.truncated = true
            break
          }
          const endRow = Math.min(startRow + batchRows - 1, scanRange.endRow)
          let mapped
          try {
            mapped = await readSheetRangeMapped(
              this.state,
              sheetId,
              {
                startRow,
                endRow,
                startColumn: scanRange.startColumn,
                endColumn: scanRange.endColumn,
              },
              meta,
              {
                signal: this.scanAbort.signal,
                // The current desktop IPC does not expose streaming byte
                // accounting; this limits accepted serialized responses after
                // receipt and before they enter the aggregate match merge.
                maxBytes: LAZY_FIND_MAX_BYTES - scannedBytes,
                isCurrent: () =>
                  this.alive && this.isLiveGeneration() && this.scanTargetIsCurrent(worksheet),
              },
            )
          } catch {
            if (!this.alive || !this.isLiveGeneration()) return
            if (this.scanAbort.signal.aborted && !this.deadlineExpired) return
            this.truncated = true
            break
          }
          if (this.deadlineExpired) {
            this.truncated = true
            break
          }
          if (!this.scanTargetIsCurrent(worksheet)) return
          if (!mapped) continue
          scannedBytes +=
            mapped.byteCount ?? new TextEncoder().encode(JSON.stringify(mapped.screen)).byteLength
          if (scannedBytes > LAZY_FIND_MAX_BYTES) {
            this.truncated = true
            break
          }
          scannedBatches += 1
          scannedCells += (endRow - startRow + 1) * scanColumns
          if (
            !mapped.raw.indexingComplete &&
            (mapped.indexedThroughScreen === null || mapped.indexedThroughScreen < endRow)
          ) {
            this.truncated = true
          }
          for (const cell of mapped.screen.cells) {
            const collectedKey = `${sheetId}|${cell.row}|${cell.column}`
            if (collectedKeys.has(collectedKey)) continue
            if (shadowed.has(`${cell.row}:${cell.column}`)) continue
            if (coveredByWindow(this.state, sheetId, cell.row, cell.column)) continue
            if (!test({ value: scalarToText(cell.value), formula: cell.formula })) continue
            collectedKeys.add(collectedKey)
            collected.push({
              row: cell.row,
              column: cell.column,
              value: cell.value,
              formula: cell.formula,
              sheetId,
            })
            if (collected.length >= LAZY_FIND_MAX_MATCHES) {
              this.truncated = true
              break
            }
          }
          this.refreshExtras(collected, comparator, unitId)
          this.emitMerged()
        }
        if (this.truncated) break
      }
      if (this.truncated) break
    }

    this.refreshExtras(collected, comparator, unitId)
    clearTimeout(this.deadlineTimer)
    if (this.truncated && this.alive && this.isLiveGeneration()) {
      // Report what was actually scanned — truncation can also come from a
      // failed read or indexing lag long before the budget.
      this.deps.setMessage(t('appFindScanTruncated', { cells: scannedCells.toLocaleString() }))
    }
    this.emitMerged()
  }

  private refreshExtras(
    collected: (ScanCell & { sheetId: string })[],
    comparator: (a: ScanCell & { sheetId: string }, b: ScanCell & { sheetId: string }) => number,
    unitId: string,
  ): void {
    const findByFormula = this.query.findBy === FindBy.FORMULA
    this.extras = [...collected]
      .sort(comparator)
      .map((cell) => makeCellMatch(unitId, cell.sheetId, cell, findByFormula))
  }
}

/// Substring replacement honoring the query's case sensitivity, replacing
/// every occurrence like Excel's Replace All.
function replaceAllOccurrences(text: string, query: IFindQuery, replaceString: string): string {
  const needle = preprocessNeedle(query)
  if (!needle) return text
  const haystack = query.caseSensitive === true ? text : text.toLowerCase()
  let result = ''
  let cursor = 0
  for (;;) {
    const index = haystack.indexOf(needle, cursor)
    if (index < 0) {
      result += text.slice(cursor)
      break
    }
    result += text.slice(cursor, index) + replaceString
    cursor = index + needle.length
  }
  return result
}

function compareTriples(a: [number, number, number], b: [number, number, number]): number {
  return a[0] - b[0] || a[1] - b[1] || a[2] - b[2]
}
