import { DrilldownSheet } from './components/DrilldownSheet'

function App() {
  return (
    <div className="box-border flex h-full min-h-screen flex-col">
      <header className="shrink-0 border-b border-gray-200 bg-neutral-50 px-5 py-3">
        <h1 className="m-0 text-xl font-semibold text-gray-900">
          Fortune Sheet — first-level drilldown
        </h1>
        <p className="mb-0 mt-1.5 text-sm text-gray-600">
          React + TypeScript +{' '}
          <code className="rounded bg-gray-100 px-1.5 py-0.5 font-mono text-xs text-gray-800">
            @fortune-sheet/react
          </code>
        </p>
      </header>
      <main className="flex min-h-0 flex-1 flex-col px-4 pb-4 pt-3">
        <DrilldownSheet />
      </main>
    </div>
  )
}

export default App
