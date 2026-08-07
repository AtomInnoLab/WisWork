import { test, expect } from '@playwright/test'
import { readFile } from 'node:fs/promises'
import {
  closeAndSaveVideo,
  createLatexE2eHarness,
  launchShell,
  openLatexProjectFromHome,
} from './helpers'

test('requires an explicit proposal apply, blocks replay, and supports undo', async () => {
  const harness = await createLatexE2eHarness('latex-ai-edit')
  const project = await harness.createProject('AI Paper')
  const launched = await launchShell({
    onboardingSeen: true,
    userDataDir: harness.userDataDir,
    videoDir: 'latex-ai-edit',
    env: { ...harness.env, WISWORK_E2E_LATEX_PROPOSAL: '1' },
  })
  try {
    const latexPage = await openLatexProjectFromHome(launched, project)
    const original = await readFile(`${project}/main.tex`, 'utf8')
    const changed = original.replace('WisWork', 'AI-confirmed WisWork')

    const review = latexPage.getByRole('region', { name: 'AI edit proposal' })
    await expect(review).toBeVisible()
    await expect(review.locator('.proposal-before')).toHaveText(original)
    await expect(review.locator('.proposal-after')).toHaveText(changed)
    const selection = review.getByRole('checkbox', { name: 'main.tex' })
    await expect(selection).toBeChecked()
    await selection.uncheck()
    await expect(review.getByRole('button', { name: 'Confirm selected changes' })).toBeDisabled()
    await selection.check()
    await review.getByRole('button', { name: 'Confirm selected changes' }).click()
    await expect(latexPage.getByRole('status').filter({ hasText: 'Changes applied' })).toBeVisible()
    await expect.poll(() => readFile(`${project}/main.tex`, 'utf8')).toBe(changed)

    await latexPage.getByRole('button', { name: 'Undo AI changes' }).click()
    await expect(
      latexPage.getByRole('status').filter({ hasText: 'AI changes were undone' }),
    ).toBeVisible()
    await expect.poll(() => readFile(`${project}/main.tex`, 'utf8')).toBe(original)

    // Keep a direct IPC subcase only for the one-shot replay security boundary.
    const proposal = await latexPage.evaluate(
      async ({ afterText }) => {
        const session = await window.latexApi.getSession()
        if (!session.ok) throw new Error('LaTeX session unavailable')
        return window.latexApi.proposeProjectEdits({
          projectId: session.value.projectId,
          files: [{ path: 'main.tex', afterText }],
        })
      },
      { afterText: changed },
    )
    expect(proposal).toMatchObject({ ok: true })

    const applied = await latexPage.evaluate(
      async ({ proposalId }) => {
        const session = await window.latexApi.getSession()
        if (!session.ok) throw new Error('LaTeX session unavailable')
        return window.latexApi.applyProposal({ projectId: session.value.projectId, proposalId })
      },
      { proposalId: proposal.ok ? proposal.value.id : '' },
    )
    expect(applied.ok).toBe(true)

    const replay = await latexPage.evaluate(
      async ({ proposalId }) => {
        const session = await window.latexApi.getSession()
        if (!session.ok) throw new Error('LaTeX session unavailable')
        return window.latexApi.applyProposal({ projectId: session.value.projectId, proposalId })
      },
      { proposalId: proposal.ok ? proposal.value.id : '' },
    )
    expect(replay).toMatchObject({ ok: false })
  } finally {
    await closeAndSaveVideo(launched, 'latex-ai-edit')
  }
})
