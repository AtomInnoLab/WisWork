import { test, expect } from '@playwright/test'
import { access, readFile } from 'node:fs/promises'
import { join } from 'node:path'
import { writeZipFixture } from '../packages/latex-project/tests/fixtures/zip'
import {
  closeAndSaveVideo,
  createLatexE2eHarness,
  editLatexSource,
  launchShell,
  openLatexProjectFromHome,
  terminateShell,
  waitForLatexPage,
} from './helpers'

test.describe('LaTeX project workflow', () => {
  test('creates, edits, saves, reports compile errors and publishes a PDF', async () => {
    const harness = await createLatexE2eHarness('latex-shell')
    const projectPath = join(harness.root, 'Created Paper')
    const launched = await launchShell({
      onboardingSeen: true,
      userDataDir: harness.userDataDir,
      videoDir: 'latex-shell',
      env: harness.env,
    })
    try {
      await launched.app.evaluate(({ dialog }, selectedPath) => {
        dialog.showSaveDialog = (async () => ({ canceled: false, filePath: selectedPath })) as never
      }, projectPath)
      await launched.page.getByTestId('quick-new-tex').click()
      const latexPage = await waitForLatexPage(launched.app)
      await expect(latexPage.getByLabel('Editor: main.tex')).toBeVisible()

      await editLatexSource(
        latexPage,
        String.raw`\documentclass{article}
\begin{document}
BROKEN
\end{document}`,
      )
      await latexPage.getByRole('button', { name: 'Compile' }).click()
      await expect(latexPage.getByRole('alert')).toHaveText('LaTeX operation failed')

      const valid = String.raw`\documentclass{article}
\begin{document}
Hello from WisWork
\end{document}`
      await editLatexSource(latexPage, valid)
      await latexPage.getByRole('button', { name: 'Compile' }).click()
      await expect(latexPage.getByText('TeX bundle ready')).toBeVisible()
      await expect.poll(() => readFile(join(projectPath, 'main.tex'), 'utf8')).toBe(valid)
      await expect(latexPage.locator('.pdf-preview canvas')).toBeVisible()
    } finally {
      await closeAndSaveVideo(launched, 'latex-shell')
    }
  })

  test('imports a regular ZIP and rejects a traversal ZIP without publishing a target', async () => {
    const harness = await createLatexE2eHarness('latex-import')
    const archive = join(harness.root, 'paper.zip')
    const target = join(harness.root, 'Imported Paper')
    await writeZipFixture(archive, [
      {
        name: 'main.tex',
        data: String.raw`\documentclass{article}\begin{document}Imported\end{document}`,
      },
      { name: 'chapter.tex', data: 'Imported chapter' },
    ])
    const launched = await launchShell({
      onboardingSeen: true,
      userDataDir: harness.userDataDir,
      videoDir: 'latex-import',
      env: harness.env,
    })
    try {
      await launched.app.evaluate(
        ({ dialog }, values) => {
          dialog.showOpenDialog = (async () => ({
            canceled: false,
            filePaths: [values.archive],
          })) as never
          dialog.showSaveDialog = (async () => ({
            canceled: false,
            filePath: values.target,
          })) as never
        },
        { archive, target },
      )
      await launched.page.getByTestId('quick-import-latex').click()
      const latexPage = await waitForLatexPage(launched.app)
      await expect(
        latexPage.getByRole('button', { name: 'chapter.tex', exact: true }),
      ).toBeVisible()
      expect(await readFile(join(target, 'chapter.tex'), 'utf8')).toBe('Imported chapter')
      await launched.page.locator('.tab-item:not(.tab-home) .tab-close').click()
      await expect(launched.page.locator('.tab-item:not(.tab-home)')).toHaveCount(0)

      const badArchive = join(harness.root, 'bad.zip')
      const badTarget = join(harness.root, 'Rejected Paper')
      await writeZipFixture(badArchive, [{ name: '../escaped.tex', data: 'escape' }])
      await launched.app.evaluate(
        ({ dialog }, values) => {
          dialog.showOpenDialog = (async () => ({
            canceled: false,
            filePaths: [values.archive],
          })) as never
          dialog.showSaveDialog = (async () => ({
            canceled: false,
            filePath: values.target,
          })) as never
        },
        { archive: badArchive, target: badTarget },
      )
      await launched.page.getByTestId('quick-import-latex').click()
      await expect(launched.page.locator('.tab-item:not(.tab-home)')).toHaveCount(0)
      await expect(access(badTarget)).rejects.toMatchObject({ code: 'ENOENT' })
      await expect(access(join(harness.root, 'escaped.tex'))).rejects.toMatchObject({
        code: 'ENOENT',
      })
    } finally {
      await closeAndSaveVideo(launched, 'latex-import')
    }
  })

  test('binds project IPC to the owning renderer and rejects paths outside the project', async () => {
    const harness = await createLatexE2eHarness('latex-ipc-boundary')
    const first = await harness.createProject('First')
    const second = await harness.createProject('Second')
    const launched = await launchShell({
      onboardingSeen: true,
      userDataDir: harness.userDataDir,
      videoDir: 'latex-ipc-boundary',
      env: harness.env,
    })
    try {
      const firstPage = await openLatexProjectFromHome(launched, first)
      const firstProjectId = await firstPage.evaluate(async () => {
        const result = await window.latexApi.getSession()
        if (!result.ok) throw new Error('LaTeX session unavailable')
        return result.value.projectId
      })
      await launched.page.locator('.tab-item:not(.tab-home) .tab-close').click()
      await expect(launched.page.locator('.tab-item:not(.tab-home)')).toHaveCount(0)
      const secondPage = await openLatexProjectFromHome(launched, second)
      const crossProject = await secondPage.evaluate(
        ({ projectId }) => window.latexApi.readFile({ projectId, path: 'main.tex' }),
        { projectId: firstProjectId },
      )
      expect(crossProject).toMatchObject({
        ok: false,
        error: { code: 'LATEX_PROJECT_SESSION_MISMATCH' },
      })
      const ownSession = await secondPage.evaluate(async () => {
        const result = await window.latexApi.getSession()
        if (!result.ok) throw new Error('LaTeX session unavailable')
        return result.value.projectId
      })
      const escaped = await secondPage.evaluate(
        ({ projectId }) => window.latexApi.readFile({ projectId, path: '../outside.tex' }),
        { projectId: ownSession },
      )
      expect(escaped).toMatchObject({ ok: false, error: { code: 'LATEX_INVALID_PAYLOAD' } })
    } finally {
      await terminateShell(launched)
    }
  })
})
