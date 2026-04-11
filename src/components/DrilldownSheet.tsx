import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
} from 'react'
import type { MouseEvent } from 'react'
import { Workbook, type WorkbookInstance } from '@fortune-sheet/react'
import type { Sheet } from '@fortune-sheet/core'
import '@fortune-sheet/react/dist/index.css'
import {
  SOURCE_DETAIL_SHEET_ID,
  SUMMARY_SHEET_ID,
  buildDrilldownCelldata,
  createInitialWorkbook,
} from '../data/mockData'

type DrilldownMode = 'same-sheet' | 'new-sheet'

/** Deep clone for sheet snapshots (plain JSON data only; never store live Immer proxies). */
function clonePlain<T>(value: T): T {
  return JSON.parse(JSON.stringify(value)) as T
}

export function DrilldownSheet() {
  const workbookRef = useRef<WorkbookInstance>(null)
  /** Default: jump to Source data tab — summary grid unchanged (Excel-style). */
  const [mode, setMode] = useState<DrilldownMode>('new-sheet')
  const [workbookData, setWorkbookData] = useState<Sheet[]>(() =>
    createInitialWorkbook(),
  )
  const [workbookKey, setWorkbookKey] = useState(0)
  const [isDrilledDown, setIsDrilledDown] = useState(false)
  const [drilldownSheetId, setDrilldownSheetId] = useState<string | null>(
    null,
  )

  const summarySnapshotRef = useRef<Sheet[] | null>(null)

  const resetDrilldownState = useCallback(() => {
    if (mode === 'same-sheet' && summarySnapshotRef.current) {
      setWorkbookData(clonePlain(summarySnapshotRef.current))
      summarySnapshotRef.current = null
      setWorkbookKey((k) => k + 1)
    } else if (mode === 'new-sheet' && drilldownSheetId) {
      const api = workbookRef.current
      if (api) {
        api.activateSheet({ id: SUMMARY_SHEET_ID })
        if (drilldownSheetId !== SOURCE_DETAIL_SHEET_ID) {
          queueMicrotask(() => {
            workbookRef.current?.deleteSheet({ id: drilldownSheetId })
          })
        }
      }
      setDrilldownSheetId(null)
    }
    setIsDrilledDown(false)
  }, [mode, drilldownSheetId])

  const runSameSheetDrilldown = useCallback(() => {
    summarySnapshotRef.current = clonePlain(workbookData)
    const detailCelldata = buildDrilldownCelldata()
    const otherSheets = clonePlain(
      workbookData.filter((s) => s.id !== SUMMARY_SHEET_ID),
    )
    otherSheets.forEach((s, i) => {
      s.order = i + 1
    })
    setWorkbookData([
      {
        id: SUMMARY_SHEET_ID,
        name: 'Detail (same sheet)',
        celldata: detailCelldata,
        row: 28,
        column: 12,
        order: 0,
      },
      ...otherSheets,
    ])
    setWorkbookKey((k) => k + 1)
    setIsDrilledDown(true)
  }, [workbookData])

  /** Excel-style: jump to the existing Source data sheet (no new tab). */
  const runNewSheetDrilldown = useCallback(() => {
    const api = workbookRef.current
    if (!api) return
    api.activateSheet({ id: SOURCE_DETAIL_SHEET_ID })
    setDrilldownSheetId(SOURCE_DETAIL_SHEET_ID)
    setIsDrilledDown(true)
  }, [])

  const handleDrilldown = useCallback(() => {
    if (isDrilledDown) return
    if (mode === 'same-sheet') {
      runSameSheetDrilldown()
    } else {
      runNewSheetDrilldown()
    }
  }, [isDrilledDown, mode, runNewSheetDrilldown, runSameSheetDrilldown])

  const handleBack = useCallback(() => {
    if (!isDrilledDown) return
    resetDrilldownState()
  }, [isDrilledDown, resetDrilldownState])

  const isDrilledDownRef = useRef(isDrilledDown)
  useEffect(() => {
    isDrilledDownRef.current = isDrilledDown
  }, [isDrilledDown])

  const actionRef = useRef({ handleDrilldown, handleBack })
  useEffect(() => {
    actionRef.current = { handleDrilldown, handleBack }
  }, [handleDrilldown, handleBack])

  /** Only the summary sheet can start a drill (detail sheet A1 is also 0,0). */
  const activeSheetIdRef = useRef<string>(SUMMARY_SHEET_ID)

  /**
   * After switching sheets, `getSelection()` is sometimes empty on `dblclick` even though
   * the second mousedown was on A1 — then we skip preventDefault and Fortune opens edit mode.
   * `afterCellMouseDown` still sees the correct cell; use it as a fallback within a short window.
   */
  const lastSummaryCellDownRef = useRef<{
    row: number
    column: number
    t: number
  } | null>(null)

  const drillTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null)
  const scheduleDrilldown = useCallback(() => {
    if (isDrilledDownRef.current) return
    if (drillTimerRef.current) window.clearTimeout(drillTimerRef.current)
    drillTimerRef.current = window.setTimeout(() => {
      drillTimerRef.current = null
      actionRef.current.handleDrilldown()
    }, 50)
  }, [])

  const selectionIncludesA1 = useCallback((api: WorkbookInstance | null) => {
    const ranges = api?.getSelection?.()
    if (!ranges?.length) return false
    return ranges.some(
      (s) =>
        s.row[0] === 0 &&
        s.row[1] === 0 &&
        s.column[0] === 0 &&
        s.column[1] === 0,
    )
  }, [])

  /** Clicks often land on selection overlays (not under `luckysheet-sheettable_*`). */
  const isFortuneGridSurface = (target: EventTarget | null) => {
    const el = target as HTMLElement | null
    if (!el?.closest) return false
    return Boolean(
      el.closest('.fortune-cell-area') || el.closest('.fortune-sheet-canvas'),
    )
  }

  /**
   * Fortune Sheet’s own double-click enters cell edit. Intercept in capture phase and
   * confirm A1 via selection and/or the last Summary mousedown from the hook.
   */
  const handleGridDoubleClickCapture = useCallback(
    (e: MouseEvent<HTMLDivElement>) => {
      if (!isFortuneGridSurface(e.target)) return
      if (activeSheetIdRef.current !== SUMMARY_SHEET_ID) return
      if (isDrilledDownRef.current) return
      const api = workbookRef.current
      const last = lastSummaryCellDownRef.current
      const recentA1 =
        last != null &&
        last.row === 0 &&
        last.column === 0 &&
        Date.now() - last.t < 650
      if (!selectionIncludesA1(api) && !recentA1) return
      e.preventDefault()
      e.stopPropagation()
      scheduleDrilldown()
    },
    [scheduleDrilldown, selectionIncludesA1],
  )

  const hooks = useMemo(
    () => ({
      afterActivateSheet: (id: string) => {
        activeSheetIdRef.current = id
        if (id !== SUMMARY_SHEET_ID) {
          lastSummaryCellDownRef.current = null
        }
      },
      afterCellMouseDown: (
        _cell: unknown,
        cellInfo: {
          row: number
          column: number
          startRow: number
          startColumn: number
          endRow: number
          endColumn: number
        },
      ) => {
        if (isDrilledDownRef.current) return
        if (activeSheetIdRef.current !== SUMMARY_SHEET_ID) {
          lastSummaryCellDownRef.current = null
          return
        }
        lastSummaryCellDownRef.current = {
          row: cellInfo.row,
          column: cellInfo.column,
          t: Date.now(),
        }
      },
    }),
    [],
  )

  const onModeChange = (next: DrilldownMode) => {
    if (next !== mode && isDrilledDown) {
      if (mode === 'same-sheet') {
        if (summarySnapshotRef.current) {
          setWorkbookData(clonePlain(summarySnapshotRef.current))
          summarySnapshotRef.current = null
          setWorkbookKey((k) => k + 1)
        }
        setIsDrilledDown(false)
      } else if (drilldownSheetId) {
        const api = workbookRef.current
        if (api) {
          api.activateSheet({ id: SUMMARY_SHEET_ID })
          if (drilldownSheetId !== SOURCE_DETAIL_SHEET_ID) {
            queueMicrotask(() => {
              workbookRef.current?.deleteSheet({ id: drilldownSheetId })
            })
          }
        }
        setDrilldownSheetId(null)
        setIsDrilledDown(false)
      }
    }
    setMode(next)
  }

  return (
    <div className="flex min-h-0 flex-1 flex-col gap-2">
      <header className="flex flex-wrap items-center gap-x-4 gap-y-3">
        <div className="flex gap-4" role="group" aria-label="Drilldown mode">
          <label className="inline-flex cursor-pointer items-center gap-1.5 text-sm text-gray-800">
            <input
              type="radio"
              name="drilldown-mode"
              checked={mode === 'new-sheet'}
              onChange={() => onModeChange('new-sheet')}
            />
            Open Source tab (Excel-style)
          </label>
          <label className="inline-flex cursor-pointer items-center gap-1.5 text-sm text-gray-800">
            <input
              type="radio"
              name="drilldown-mode"
              checked={mode === 'same-sheet'}
              onChange={() => onModeChange('same-sheet')}
            />
            Same sheet (demo)
          </label>
        </div>
        {isDrilledDown && (
          <button
            type="button"
            className="cursor-pointer rounded-md border border-blue-600 bg-blue-600 px-3.5 py-1.5 text-sm text-white hover:bg-blue-700"
            onClick={handleBack}
          >
            Back to summary
          </button>
        )}
      </header>
      <p className="m-0 text-sm text-gray-700">
        <strong>Double-click</strong> cell <strong>A1</strong> on the <strong>Summary</strong> sheet (like Excel Pivot “Show Details”; see{' '}
        <a
          className="text-blue-700 underline"
          href="https://www.youtube.com/watch?v=jYQ8Bj3ki0g"
          target="_blank"
          rel="noreferrer"
        >
          this walkthrough
        </a>
        ). The <strong>Source data</strong> tab holds row-level rows that roll up to the summary.
        {mode === 'new-sheet'
          ? ' Drill switches to the Source data tab (summary unchanged). Use Back to return to Summary.'
          : ' Demo mode replaces the Summary grid in place; Source data stays; use Back to restore.'}
      </p>
      <div
        className="flex h-full min-h-[480px] flex-1 overflow-hidden rounded-lg border border-gray-200 bg-white [&_.fortune-container]:h-full [&_.fortune-container]:min-h-[460px]"
        onDoubleClickCapture={handleGridDoubleClickCapture}
      >
        <Workbook
          key={workbookKey}
          ref={workbookRef}
          data={workbookData}
          hooks={hooks}
          lang="en"
          showToolbar
          showFormulaBar
          showSheetTabs
        />
      </div>
    </div>
  )
}
