import { spawnSync } from 'node:child_process'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

export const sevenHostGoldens = Object.freeze([
  {
    host: 'latex',
    files: [
      'apps/latex/tests/agent-runtime-loop.test.ts',
      'apps/latex/tests/proposal-workflow.test.ts',
      'apps/latex/tests/enhanced-production-golden.test.ts',
    ],
  },
  {
    host: 'slides',
    files: [
      'apps/slides/tests/agent-controller.test.ts',
      'apps/slides/tests/presentation-transaction-desktop-host.test.ts',
      'apps/slides/tests/enhanced-production-golden.test.ts',
    ],
  },
  {
    host: 'docs',
    files: [
      'apps/docs/tests/agent-controller.test.ts',
      'apps/docs/tests/enhanced-mutation-confirmation.test.ts',
      'apps/docs/tests/enhanced-production-golden.test.ts',
    ],
  },
  {
    host: 'sheets',
    files: [
      'apps/sheets/tests/agent-controller.test.ts',
      'apps/sheets/tests/workbook-transactions.test.ts',
      'apps/sheets/tests/enhanced-production-golden.test.ts',
    ],
  },
  {
    host: 'office-word',
    files: [
      'packages/office-bridge/tests/seven-host-golden.test.ts',
      'apps/office-addin/tests/elevated-browser-adapters.test.ts',
    ],
    testName: 'office-word Enhanced bridge|Word compiles',
  },
  {
    host: 'office-excel',
    files: [
      'packages/office-bridge/tests/seven-host-golden.test.ts',
      'apps/office-addin/tests/elevated-browser-adapters.test.ts',
    ],
    testName: 'office-excel Enhanced bridge|Excel compiles',
  },
  {
    host: 'office-powerpoint',
    files: [
      'packages/office-bridge/tests/seven-host-golden.test.ts',
      'apps/office-addin/tests/elevated-browser-adapters.test.ts',
    ],
    testName: 'office-powerpoint Enhanced bridge|PowerPoint compiles',
  },
])

export function runSevenHostGoldens(root = join(dirname(fileURLToPath(import.meta.url)), '..')) {
  for (const golden of sevenHostGoldens) {
    const args = [
      join(root, 'node_modules/vitest/vitest.mjs'),
      'run',
      ...golden.files,
      '--reporter=verbose',
      '--disableConsoleIntercept',
    ]
    if (golden.testName) args.push('-t', golden.testName)
    const result = spawnSync(process.execPath, args, { cwd: root, encoding: 'utf8' })
    process.stdout.write(result.stdout ?? '')
    process.stderr.write(result.stderr ?? '')
    if (result.status !== 0) return result.status ?? 1
    const reports = [...(result.stdout ?? '').matchAll(/ENHANCED_GOLDEN_REPORT (\{[^\n]+\})/g)]
      .map((match) => JSON.parse(match[1]))
      .filter((report) => report.host === golden.host)
    if (
      reports.length !== 1 ||
      Object.keys(reports[0]).sort().join(',') !== 'host,rollback,verification' ||
      reports[0].verification !== 'verified' ||
      reports[0].rollback !== 'restored'
    ) {
      console.error(`Invalid Enhanced golden report for ${golden.host}`)
      return 1
    }
  }
  return 0
}

if (process.argv[2] === 'execute-goldens') process.exit(runSevenHostGoldens())
