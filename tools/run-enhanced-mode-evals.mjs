import { spawnSync } from 'node:child_process'
import { existsSync, readFileSync } from 'node:fs'
import { dirname, join, relative, sep } from 'node:path'
import { fileURLToPath } from 'node:url'

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..')
const readJson = (path) => JSON.parse(readFileSync(join(ROOT, path), 'utf8'))

const WORKSPACES = [
  { prefix: 'apps/shell/', workspace: '@wiswork/shell' },
  { prefix: 'apps/latex/', workspace: '@wiswork/latex' },
  { prefix: 'packages/latex-project/', workspace: '@wiswork/latex-project' },
  { prefix: 'packages/codex-bridge/', workspace: '@wiswork/codex-bridge' },
]

function evidenceCommand(path) {
  const match = WORKSPACES.find(({ prefix }) => path.startsWith(prefix))
  if (!match || !path.endsWith('.test.ts') || !existsSync(join(ROOT, path))) {
    throw new Error(`enhanced_mode_eval_evidence_invalid:${path}`)
  }
  const testPath = relative(join(ROOT, match.prefix), join(ROOT, path)).split(sep).join('/')
  if (testPath.startsWith('../')) throw new Error(`enhanced_mode_eval_evidence_invalid:${path}`)
  return { workspace: match.workspace, testPath }
}

const evaluation = readJson('packages/codex-bridge/evals/latex-runtime-pilot.json')
const security = readJson('packages/codex-bridge/evals/security-matrix.json')
const evidencePaths = [
  ...evaluation.cases.map((item) => item.automatedEvidence),
  ...security.cases.map((item) => item.evidence),
]
const uniqueEvidence = [...new Set(evidencePaths)]
const outcomes = new Map()

for (const path of uniqueEvidence) {
  const { workspace, testPath } = evidenceCommand(path)
  const result = spawnSync(
    process.platform === 'win32' ? 'npm.cmd' : 'npm',
    ['run', 'test', '-w', workspace, '--', '--run', testPath],
    { cwd: ROOT, stdio: 'inherit', env: process.env },
  )
  outcomes.set(path, result.status === 0)
}

const enhancedPassRate =
  evaluation.cases.filter((item) => outcomes.get(item.automatedEvidence) === true).length /
  evaluation.cases.length
const parityRate =
  evaluation.cases.filter(
    (item) => outcomes.get(item.automatedEvidence) === (item.legacy === 'pass'),
  ).length / evaluation.cases.length
const securityPassRate =
  security.cases.filter((item) => outcomes.get(item.evidence) === true).length /
  security.cases.length

if (
  enhancedPassRate < evaluation.thresholds.minimumScenarioPassRate ||
  parityRate < evaluation.thresholds.minimumLegacyParityRate ||
  securityPassRate < security.requiredPassRate
) {
  throw new Error('enhanced_mode_eval_threshold_failed')
}

process.stdout.write(
  `enhanced_mode_evals_ok scenarios=${enhancedPassRate.toFixed(2)} parity=${parityRate.toFixed(2)} security=${securityPassRate.toFixed(2)}\n`,
)
