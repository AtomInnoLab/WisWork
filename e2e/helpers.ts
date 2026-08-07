/**
 * Shared launcher for Electron E2E tests.
 *
 * Each test boots the built shell (`apps/shell/out`) against a scratch
 * userData dir (via WISWORK_USER_DATA) so runs never touch real settings
 * and never collide with a running install's single-instance lock.
 * Build first: `npm run build:all`.
 */
import { _electron as electron, type ElectronApplication, type Page } from '@playwright/test'
import { createHash } from 'node:crypto'
import { chmod, mkdir, mkdtemp, writeFile } from 'node:fs/promises'
import { existsSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { createRequire } from 'node:module'
import type { LatexApi } from '../apps/latex/src/shared/ipc'

declare global {
  interface Window {
    latexApi: LatexApi
  }
}

export const SHELL_DIR = resolve(__dirname, '../apps/shell')
export const ARTIFACTS_DIR = resolve(__dirname, 'artifacts')

const SHELL_MAIN = join(SHELL_DIR, 'out/main/index.js')

interface LaunchOptions {
  /** reuse a previous scratch dir to simulate a second launch */
  userDataDir?: string
  /** UI language override (WISWORK_LANG); defaults to English for stable assertions */
  lang?: string
  /** pre-seed app-settings.json with onboardingSeen=true to start at the home screen */
  onboardingSeen?: boolean
  /** subdir of e2e/artifacts to store this launch's video in */
  videoDir: string
  /** absolute document path passed as argv, opened in an editor tab on launch */
  openFile?: string
  /** additional deterministic fixture settings for this launch */
  env?: Readonly<Record<string, string>>
}

export interface LaunchedApp {
  app: ElectronApplication
  page: Page
  userDataDir: string
}

export async function launchShell(options: LaunchOptions): Promise<LaunchedApp> {
  if (!existsSync(SHELL_MAIN)) {
    throw new Error(`Missing build output at ${SHELL_MAIN} — run \`npm run build:all\` first`)
  }
  const userDataDir = options.userDataDir ?? (await mkdtemp(join(tmpdir(), 'wiswork-e2e-')))
  await mkdir(userDataDir, { recursive: true })
  if (options.onboardingSeen) {
    await writeFile(
      join(userDataDir, 'app-settings.json'),
      JSON.stringify({ onboardingSeen: true }),
    )
  }
  const require = createRequire(join(SHELL_DIR, 'package.json'))
  const executablePath = require('electron') as unknown as string
  // ELECTRON_RUN_AS_NODE (set by VS Code/CI hosts) would boot Electron as
  // plain Node with no windows — strip it so the app always starts as an app
  const { ELECTRON_RUN_AS_NODE: _electronRunAsNode, ...hostEnv } = process.env
  // Linux CI runners restrict unprivileged user namespaces (no usable SUID
  // sandbox) and run under xvfb without GPU — without these the window opens
  // but the renderer never loads. The suite drives trusted local builds only.
  // Switches go before the app path so Chromium is guaranteed to consume them
  // and they never leak into the argv the app parses for documents to open.
  const args: string[] = []
  if (process.platform === 'linux') args.push('--no-sandbox', '--disable-gpu')
  args.push(SHELL_DIR)
  if (options.openFile) args.push(options.openFile)
  const app = await electron.launch({
    executablePath,
    args,
    env: {
      ...hostEnv,
      WISWORK_USER_DATA: userDataDir,
      WISWORK_LANG: options.lang ?? 'en',
      ...options.env,
      ...(process.platform === 'linux' ? { ELECTRON_DISABLE_SANDBOX: '1' } : {}),
    },
    // Playwright's Electron screencast wedges the page CDP session on Linux
    // (page.url() stays empty, no lifecycle events, evaluate hangs) — record
    // only where it works
    recordVideo:
      process.platform === 'linux'
        ? undefined
        : {
            dir: join(ARTIFACTS_DIR, options.videoDir),
            size: { width: 1280, height: 800 },
          },
  })
  const page = await app.firstWindow()
  await waitForDocumentReady(app, page)
  return { app, page, userDataDir }
}

export interface LatexE2eHarness {
  root: string
  userDataDir: string
  env: Readonly<Record<string, string>>
  createProject(name: string, source?: string): Promise<string>
}

/**
 * Real Electron/IPC LaTeX harness with deterministic local external assets.
 * The fake executable stands in only for Tectonic itself; project access,
 * bundle verification, workspace isolation, publishing and UI are production code.
 */
export async function createLatexE2eHarness(name: string): Promise<LatexE2eHarness> {
  const root = await mkdtemp(join(tmpdir(), `wiswork-${name}-`))
  const userDataDir = join(root, 'user-data')
  await mkdir(userDataDir, { recursive: true })

  const bundle = Buffer.from('wiswork-e2e-tectonic-bundle\n')
  const bundleId = 'wiswork-e2e-bundle'
  const bundleDirectory = join(userDataDir, 'latex', 'bundles')
  await mkdir(bundleDirectory, { recursive: true })
  await writeFile(join(bundleDirectory, `${bundleId}.tar`), bundle)

  const pdfPath = join(root, 'success.pdf')
  await writeFile(pdfPath, minimalPdf('WisWork LaTeX E2E'))
  const tectonicPath = join(root, 'tectonic-e2e')
  await writeFile(
    tectonicPath,
    `#!/bin/bash
main="$1"
shift
out=""
while [ "$#" -gt 0 ]; do
  if [ "$1" = "--outdir" ]; then
    shift
    out="$1"
  fi
  shift
done
if /bin/grep -q BROKEN "$main"; then
  echo "error: $main:3: forced E2E syntax error" >&2
  exit 1
fi
stem="\${main##*/}"
stem="\${stem%.*}"
/bin/cp ${shellQuote(pdfPath)} "$out/$stem.pdf"
echo "E2E compile succeeded"
`,
  )
  await chmod(tectonicPath, 0o700)

  const manifestPath = join(root, 'tectonic-manifest.json')
  await writeFile(
    manifestPath,
    JSON.stringify({
      schemaVersion: 1,
      tectonic: {
        version: '0.16.9',
        license: {
          spdx: 'MIT',
          sourceUrl: 'https://github.com/tectonic-typesetting/tectonic/blob/66b6654/LICENSE',
        },
        assets: [
          {
            id: 'wiswork-e2e-tectonic',
            platform: 'e2e-linux',
            url: 'https://github.com/tectonic-typesetting/tectonic/releases/download/tectonic%400.16.9/tectonic-0.16.9-x86_64-unknown-linux-gnu.tar.gz',
            bytes: 1,
            sha256: '0'.repeat(64),
            archive: { format: 'tar.gz', executable: 'tectonic' },
          },
        ],
      },
      bundle: {
        id: bundleId,
        url: 'https://relay.fullyjustified.net/default_bundle_v33.tar',
        bytes: bundle.byteLength,
        sha256: createHash('sha256').update(bundle).digest('hex'),
        license: {
          spdx: 'LicenseRef-TeX-Live-Collection',
          sourceUrl: 'https://tug.org/texlive/copying.html',
        },
      },
    }),
  )

  return {
    root,
    userDataDir,
    env: {
      WISWORK_TECTONIC_PATH: tectonicPath,
      WISWORK_TECTONIC_MANIFEST_PATH: manifestPath,
    },
    async createProject(projectName, source = defaultLatexSource()) {
      const project = join(root, projectName)
      await mkdir(project)
      await writeFile(join(project, 'main.tex'), source)
      return project
    },
  }
}

export async function waitForLatexPage(app: ElectronApplication): Promise<Page> {
  return waitForPageWithUrl(app, 'latex/out/renderer')
}

export async function openLatexProjectFromHome(
  launched: LaunchedApp,
  projectPath: string,
): Promise<Page> {
  const before = new Set(launched.app.windows())
  await launched.page.evaluate(
    (path) =>
      (
        window as unknown as {
          aiOffice: { openLatexProject(project: string): Promise<unknown> }
        }
      ).aiOffice.openLatexProject(path),
    projectPath,
  )
  const deadline = Date.now() + 30_000
  let page: Page | undefined
  while (!page && Date.now() < deadline) {
    for (const candidate of launched.app.windows()) {
      if (before.has(candidate)) continue
      const href = await candidate.evaluate(() => window.location.href).catch(() => candidate.url())
      if (href.includes('latex/out/renderer')) {
        page = candidate
        break
      }
    }
    if (!page) await launched.app.waitForEvent('window', { timeout: 500 }).catch(() => undefined)
  }
  if (!page) throw new Error('A new LaTeX renderer was not created')
  if (!before.has(page)) {
    try {
      await page.locator('.latex-workbench').waitFor({ state: 'visible', timeout: 10_000 })
    } catch (error) {
      const rendererErrors: string[] = []
      page.on('pageerror', (reason) => rendererErrors.push(reason.message))
      await page.reload().catch(() => undefined)
      await page.waitForTimeout(500)
      const diagnostics = await page
        .evaluate(() => ({
          url: window.location.href,
          body: document.body.innerText,
          html: document.body.innerHTML,
          hasApi: typeof window.latexApi !== 'undefined',
          scripts: Array.from(document.scripts, (script) => script.src),
        }))
        .catch(() => ({ url: page.url(), body: '<renderer unavailable>' }))
      throw new Error(
        `LaTeX renderer did not mount: ${JSON.stringify({ ...diagnostics, rendererErrors })}`,
        {
          cause: error,
        },
      )
    }
  }
  return page
}

export async function editLatexSource(page: Page, source: string): Promise<void> {
  const editor = page.locator('.latex-editor .cm-content')
  await editor.fill(source)
  await page.waitForTimeout(50)
}

function defaultLatexSource(): string {
  return String.raw`\documentclass{article}
\begin{document}
WisWork
\end{document}`
}

function shellQuote(value: string): string {
  return `'${value.replaceAll("'", `'\\''`)}'`
}

function minimalPdf(text: string): Buffer {
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 144] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>',
    `<< /Length ${text.length + 31} >>\nstream\nBT /F1 18 Tf 36 72 Td (${text}) Tj ET\nendstream`,
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
  ]
  let body = '%PDF-1.4\n'
  const offsets = [0]
  for (let index = 0; index < objects.length; index += 1) {
    offsets.push(Buffer.byteLength(body))
    body += `${index + 1} 0 obj\n${objects[index]}\nendobj\n`
  }
  const xref = Buffer.byteLength(body)
  body += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`
  body += offsets
    .slice(1)
    .map((offset) => `${String(offset).padStart(10, '0')} 00000 n \n`)
    .join('')
  body += `trailer << /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`
  return Buffer.from(body)
}

/**
 * Playwright can attach to the Electron window mid-navigation and miss the
 * load lifecycle events entirely (Linux timing) — waitForLoadState then hangs
 * on a page that is actually loaded. Polling through evaluate uses the live
 * CDP session instead of the missed events.
 */
async function waitForDocumentReady(
  app: ElectronApplication,
  page: Page,
  timeoutMs = 30_000,
): Promise<void> {
  const deadline = Date.now() + timeoutMs
  while (Date.now() < deadline) {
    // the pre-navigation about:blank document also reports readyState complete
    const ready = await page
      .evaluate(() =>
        document.readyState !== 'loading' && window.location.href !== 'about:blank'
          ? document.readyState
          : null,
      )
      .catch(() => null)
    if (ready) return
    await new Promise((r) => setTimeout(r, 100))
  }
  // process list tells renderer-spawn failures apart from slow loads
  const diag = await app
    .evaluate(({ app: electronApp, BrowserWindow }) => ({
      processes: electronApp.getAppMetrics().map((m) => m.type),
      contents: BrowserWindow.getAllWindows().map((w) => w.webContents.getURL()),
    }))
    .catch((e) => String(e))
  throw new Error(`Shell window never loaded (url: ${page.url()}, diag: ${JSON.stringify(diag)})`)
}

/**
 * Close the app and return the recorded video path for the given page.
 *
 * Open editor tabs trigger a native Save/Don't Save/Cancel dialog on close,
 * which would block app.close() forever — stub the dialog to answer
 * "Don't Save" (button index 1) so shutdown stays unattended. If close still
 * hangs, kill the process after 20s so the suite never wedges.
 */
export async function closeAndSaveVideo(
  launched: LaunchedApp,
  name: string,
): Promise<string | undefined> {
  const video = launched.page.video()
  await launched.app
    .evaluate(({ dialog }) => {
      dialog.showMessageBox = (async () => ({
        response: 1,
        checkboxChecked: false,
      })) as typeof dialog.showMessageBox
    })
    .catch(() => {})
  let killTimer: NodeJS.Timeout | undefined
  await Promise.race([
    launched.app.close(),
    new Promise<void>((resolvePromise) => {
      killTimer = setTimeout(() => {
        launched.app.process().kill()
        resolvePromise()
      }, 20_000)
    }),
  ])
  if (killTimer) clearTimeout(killTimer)
  if (!video) return undefined
  const target = join(ARTIFACTS_DIR, 'videos', `${name}.webm`)
  try {
    await video.saveAs(target)
    return target
  } catch {
    return undefined
  }
}

export function screenshotPath(name: string): string {
  return join(ARTIFACTS_DIR, 'screenshots', `${name}.png`)
}

/** Force-stop only for restart/multi-view cases that reproduce Electron's native shutdown wedge. */
export async function terminateShell(launched: LaunchedApp): Promise<void> {
  const closed = launched.app.waitForEvent('close', { timeout: 10_000 }).catch(() => undefined)
  launched.app.process().kill('SIGKILL')
  await closed
}

/**
 * Wait for a page whose URL contains `urlPart` (e.g. an editor WebContentsView).
 * Checks windows that already exist before listening, so it never races the
 * view being created between launch and the first waitForEvent call.
 */
export async function waitForPageWithUrl(
  app: ElectronApplication,
  urlPart: string,
  timeoutMs = 30_000,
): Promise<Page> {
  const deadline = Date.now() + timeoutMs
  for (;;) {
    for (const candidate of app.windows()) {
      if (candidate.url().includes(urlPart)) return candidate
      // page.url() stays empty when attach raced navigation; ask the document
      const href = await candidate.evaluate(() => window.location.href).catch(() => '')
      if (href.includes(urlPart)) return candidate
    }
    const remaining = deadline - Date.now()
    if (remaining <= 0) throw new Error(`No window with URL containing "${urlPart}"`)
    await app.waitForEvent('window', { timeout: Math.min(remaining, 1_000) }).catch(() => {})
  }
}
