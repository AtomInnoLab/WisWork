import { test, expect } from '@playwright/test'
import { readFile } from 'node:fs/promises'
import {
  createLatexE2eHarness,
  launchShell,
  openLatexProjectFromHome,
  terminateShell,
  waitForLatexPage,
} from './helpers'

test('restores valid LaTeX tabs and the active project after restart', async () => {
  const harness = await createLatexE2eHarness('latex-tab-restore')
  const first = await harness.createProject('Restore One')
  const second = await harness.createProject('Restore Two')
  const launched = await launchShell({
    onboardingSeen: true,
    userDataDir: harness.userDataDir,
    videoDir: 'latex-tab-restore-first',
    env: harness.env,
  })
  await openLatexProjectFromHome(launched, first)
  await openLatexProjectFromHome(launched, second)
  await expect
    .poll(async () => {
      const stored = JSON.parse(
        await readFile(`${harness.userDataDir}/open-tabs.json`, 'utf8'),
      ) as {
        projectPaths: string[]
      }
      return stored.projectPaths.length
    })
    .toBe(2)
  await terminateShell(launched)

  const restored = await launchShell({
    onboardingSeen: true,
    userDataDir: harness.userDataDir,
    videoDir: 'latex-tab-restore-second',
    env: harness.env,
  })
  try {
    await expect(restored.page.locator('.tab-item:not(.tab-home)')).toHaveCount(2)
    await expect(restored.page.locator('.tab-item.active .tab-title')).toHaveText('Restore Two')
    const latexPage = await waitForLatexPage(restored.app)
    await expect(latexPage.getByLabel('Editor: main.tex')).toBeVisible()
  } finally {
    await terminateShell(restored)
  }
})
