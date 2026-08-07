import { test, expect } from '@playwright/test'
import { readFile } from 'node:fs/promises'
import {
  closeAndSaveVideo,
  createLatexE2eHarness,
  editLatexSource,
  launchShell,
  openLatexProjectFromHome,
} from './helpers'

test('dirty LaTeX close can be cancelled, then discarded explicitly', async () => {
  const harness = await createLatexE2eHarness('latex-dirty-close')
  const project = await harness.createProject('Dirty Paper')
  const launched = await launchShell({
    onboardingSeen: true,
    userDataDir: harness.userDataDir,
    videoDir: 'latex-dirty-close',
    env: harness.env,
  })
  try {
    const latexPage = await openLatexProjectFromHome(launched, project)
    const original = await readFile(`${project}/main.tex`, 'utf8')
    const edited = String.raw`\documentclass{article}\begin{document}Dirty edit\end{document}`
    await editLatexSource(latexPage, edited)

    await launched.app.evaluate(({ dialog }) => {
      dialog.showMessageBox = (async () => ({ response: 2, checkboxChecked: false })) as never
    })
    await launched.page.locator('.tab-item:not(.tab-home) .tab-close').click()
    await expect(launched.page.locator('.tab-item:not(.tab-home)')).toHaveCount(1)

    await launched.app.evaluate(({ dialog }) => {
      dialog.showMessageBox = (async () => ({ response: 1, checkboxChecked: false })) as never
    })
    await launched.page.locator('.tab-item:not(.tab-home) .tab-close').click()
    await expect(launched.page.locator('.tab-item:not(.tab-home)')).toHaveCount(0)
    expect(await readFile(`${project}/main.tex`, 'utf8')).toBe(original)
  } finally {
    await closeAndSaveVideo(launched, 'latex-dirty-close')
  }
})
