import { test, expect } from '@playwright/test'
import { access, readFile, writeFile } from 'node:fs/promises'
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
      let latexPage = await waitForLatexPage(launched.app)
      await expect(latexPage.getByLabel('Editor: main.tex')).toBeVisible()
      await expect(latexPage.locator('.ai-panel-title').getByText('WisWork AI')).toBeVisible()
      await expect(latexPage.getByRole('tab', { name: '编译' })).toHaveCount(0)
      const toolbar = latexPage.getByRole('toolbar', { name: 'LaTeX toolbar' })

      await editLatexSource(
        latexPage,
        String.raw`\documentclass{article}
\begin{document}
BROKEN
\end{document}`,
      )
      await toolbar.getByRole('tab', { name: 'Compile', exact: true }).click()
      await toolbar.getByRole('button', { name: 'Compile', exact: true }).click()
      await expect(latexPage.getByRole('alert')).toHaveText('LaTeX operation failed')

      const valid = String.raw`\documentclass{article}
\begin{document}
Hello from WisWork
\end{document}`
      await editLatexSource(latexPage, valid)
      await toolbar.getByRole('button', { name: 'Compile', exact: true }).click()
      await toolbar.getByRole('button', { name: /Problems/ }).click()
      await expect(latexPage.getByText('Remote TeX bundle configured')).toBeVisible()
      await expect.poll(() => readFile(join(projectPath, 'main.tex'), 'utf8')).toBe(valid)
      await expect(latexPage.locator('.pdf-preview canvas')).toBeVisible()

      await expect
        .poll(async () => {
          const boxes = await Promise.all(
            ['.project-tree header', '.open-tabs', '.readonly-pdf-toolbar', '.ai-panel-header'].map(
              (selector) => latexPage.locator(selector).boundingBox(),
            ),
          )
          return boxes.map((box) => ({ y: Math.round(box?.y ?? -1), height: box?.height ?? -1 }))
        })
        .toEqual([
          { y: 120, height: 48 },
          { y: 120, height: 48 },
          { y: 120, height: 48 },
          { y: 120, height: 48 },
        ])

      await latexPage.getByRole('button', { name: 'Collapse AI panel' }).click()
      const expandAi = latexPage.getByRole('button', { name: 'Expand AI panel' })
      await expect(expandAi).toBeVisible()
      await expandAi.click()
      await expect(latexPage.getByRole('button', { name: 'Collapse AI panel' })).toBeVisible()

      const exportedPdf = join(harness.root, 'exported.pdf')
      await writeFile(join(projectPath, 'unopened-dependency.tex'), 'changed after compile')
      await launched.app.evaluate(({ dialog }, selectedPath) => {
        dialog.showSaveDialog = (async () => ({ canceled: false, filePath: selectedPath })) as never
      }, exportedPdf)
      await toolbar.getByRole('tab', { name: 'PDF', exact: true }).click()
      await toolbar.getByRole('button', { name: 'Export PDF', exact: true }).click()
      const staleDialog = latexPage.getByRole('dialog', { name: 'PDF preview is out of date' })
      await expect(staleDialog).toBeVisible()
      await staleDialog.getByRole('button', { name: 'Export last PDF', exact: true }).click()
      await expect(staleDialog).toBeHidden()
      await expect.poll(() => readFile(exportedPdf, 'utf8')).toContain('%PDF')

      await latexPage.getByRole('button', { name: 'Close PDF preview' }).click()
      await expect(latexPage.locator('.pdf-preview')).toHaveCount(0)
      await latexPage.getByRole('button', { name: 'Open PDF preview' }).click()
      await expect(latexPage.locator('.pdf-preview canvas')).toBeVisible()

      await launched.page.locator('.tab-item:not(.tab-home) .tab-close').click()
      await expect(launched.page.locator('.tab-item:not(.tab-home)')).toHaveCount(0)
      latexPage = await openLatexProjectFromHome(launched, projectPath)
      await expect(latexPage.locator('.pdf-preview canvas')).toBeVisible()
    } finally {
      await closeAndSaveVideo(launched, 'latex-shell')
    }
  })

  test('creates, renames, and confirms deletion from the project tree', async () => {
    const harness = await createLatexE2eHarness('latex-file-actions')
    const project = await harness.createProject('File Actions')
    const launched = await launchShell({
      onboardingSeen: true,
      userDataDir: harness.userDataDir,
      videoDir: 'latex-file-actions',
      env: harness.env,
    })
    try {
      const latexPage = await openLatexProjectFromHome(launched, project)
      await latexPage.getByTitle('New file').click()
      const pathInput = latexPage.getByLabel('Project-relative path')
      await pathInput.fill('chapter.tex')
      await latexPage.getByRole('button', { name: 'Create', exact: true }).click()
      const projectTree = latexPage.locator('.project-tree')
      const chapter = projectTree.getByRole('button', { name: 'chapter.tex', exact: true })
      await expect(chapter).toBeVisible()
      await expect(access(join(project, 'chapter.tex'))).resolves.toBeUndefined()

      await chapter.hover()
      await latexPage.getByRole('button', { name: 'File actions for chapter.tex' }).click()
      await latexPage.getByRole('menuitem', { name: 'Rename' }).click()
      await pathInput.fill('renamed.tex')
      await latexPage.getByRole('button', { name: 'Rename', exact: true }).click()
      const renamed = projectTree.getByRole('button', { name: 'renamed.tex', exact: true })
      await expect(renamed).toBeVisible()

      await renamed.hover()
      await latexPage.getByRole('button', { name: 'File actions for renamed.tex' }).click()
      await latexPage.getByRole('menuitem', { name: 'Delete' }).click()
      await expect(latexPage.getByRole('dialog')).toContainText('cannot be undone')
      await latexPage.getByRole('button', { name: 'Delete', exact: true }).click()
      await expect(latexPage.getByRole('dialog')).toBeHidden()
      await expect
        .poll(async () =>
          access(join(project, 'renamed.tex')).then(
            () => true,
            () => false,
          ),
        )
        .toBe(false)
    } finally {
      await closeAndSaveVideo(launched, 'latex-file-actions')
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
