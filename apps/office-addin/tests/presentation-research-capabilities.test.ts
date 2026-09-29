import { expect, it, vi } from 'vitest'
import { readPresentationResearchCapabilities } from '../src/skills/powerpoint/presentation-research-capabilities'
it('negotiates research cleanup and version two explicitly', async () => {
  const request = vi.fn(async () => ({
    version: 1,
    available: true,
    cleanupAvailable: true,
    historyVersions: [1, 2],
  }))
  expect(await readPresentationResearchCapabilities(request)).toEqual({
    available: true,
    cleanupAvailable: true,
    historyVersion: 2,
  })
  expect(request).toHaveBeenCalledExactlyOnceWith('research_capabilities', { includeCleanup: true })
})
it('preserves a legacy capability response without sending new history fields', async () => {
  const request = vi.fn(async () => ({ version: 1, available: true }))
  expect(await readPresentationResearchCapabilities(request)).toEqual({
    available: true,
    cleanupAvailable: false,
  })
  expect(request).toHaveBeenCalledTimes(1)
})
it('falls back once for a legacy server rejecting the extended capability request', async () => {
  const request = vi
    .fn()
    .mockRejectedValueOnce(new Error('presentation_upgrade_required'))
    .mockResolvedValueOnce({ version: 1, available: true })
  expect(await readPresentationResearchCapabilities(request)).toEqual({
    available: true,
    cleanupAvailable: false,
  })
  expect(request.mock.calls).toEqual([
    ['research_capabilities', { includeCleanup: true }],
    ['research_capabilities', {}],
  ])
})
it.each([
  { version: 2, available: true },
  { version: 1, available: true, cleanupAvailable: true },
  { version: 1, available: true, cleanupAvailable: true, historyVersions: [1, 1, 2] },
  {
    version: 1,
    available: true,
    cleanupAvailable: true,
    historyVersions: [1, 2],
    raw: 'not metadata',
  },
  { version: 1, available: true, cleanupAvailable: false, historyVersions: [1, 2] },
])('rejects malformed capabilities rather than enabling deletion: %j', async (value) => {
  const request = vi.fn(async () => value)
  await expect(readPresentationResearchCapabilities(request)).rejects.toThrow(
    'presentation_response_invalid',
  )
  expect(request).toHaveBeenCalledTimes(1)
})
it('does not retry network failure or repeatedly retry an unavailable endpoint', async () => {
  const failed = vi.fn(async () => {
    throw new Error('presentation_service_unavailable')
  })
  await expect(readPresentationResearchCapabilities(failed)).rejects.toThrow(
    'presentation_service_unavailable',
  )
  expect(failed).toHaveBeenCalledTimes(1)
  const missing = vi.fn(async () => {
    throw new Error('presentation_upgrade_required')
  })
  await expect(readPresentationResearchCapabilities(missing)).rejects.toThrow(
    'presentation_upgrade_required',
  )
  expect(missing).toHaveBeenCalledTimes(2)
})
it('opts into recovery and accepts its exact additional field only when opted in', async () => {
  const response = {
      version: 1,
      available: true,
      cleanupAvailable: true,
      recoveryAvailable: true,
      historyVersions: [1, 2],
    },
    request = vi.fn(async () => response)
  expect(await readPresentationResearchCapabilities(request, { includeRecovery: true })).toEqual({
    available: true,
    cleanupAvailable: true,
    historyVersion: 2,
    recoveryAvailable: true,
  })
  expect(request).toHaveBeenCalledExactlyOnceWith('research_capabilities', {
    includeCleanup: true,
    includeRecovery: true,
  })
  await expect(readPresentationResearchCapabilities(request)).rejects.toThrow(
    'presentation_response_invalid',
  )
})
it('uses recovery to cleanup to legacy fallback at most three times', async () => {
  const request = vi
    .fn()
    .mockRejectedValueOnce(new Error('presentation_upgrade_required'))
    .mockRejectedValueOnce(new Error('presentation_upgrade_required'))
    .mockResolvedValueOnce({ version: 1, available: true })
  expect(await readPresentationResearchCapabilities(request, { includeRecovery: true })).toEqual({
    available: true,
    cleanupAvailable: false,
  })
  expect(request.mock.calls).toEqual([
    ['research_capabilities', { includeCleanup: true, includeRecovery: true }],
    ['research_capabilities', { includeCleanup: true }],
    ['research_capabilities', {}],
  ])
})
it('accepts cleanup-only or legacy replies to an opt-in request without inventing recovery false', async () => {
  for (const response of [
    { version: 1, available: true, cleanupAvailable: true, historyVersions: [1, 2] },
    { version: 1, available: true },
  ]) {
    const request = vi.fn(async () => response),
      value = await readPresentationResearchCapabilities(request, { includeRecovery: true })
    expect(value).not.toHaveProperty('recoveryAvailable')
    expect(request).toHaveBeenCalledTimes(1)
  }
})
it('does not downgrade recovery network/response errors and stops after three unsupported reads', async () => {
  for (const response of [
    {
      version: 1,
      available: true,
      cleanupAvailable: true,
      recoveryAvailable: false,
      historyVersions: [1, 2],
    },
    {
      version: 1,
      available: true,
      cleanupAvailable: true,
      recoveryAvailable: true,
      historyVersions: [2, 1],
    },
    {
      version: 1,
      available: true,
      cleanupAvailable: true,
      recoveryAvailable: true,
      historyVersions: [1, 2],
      unknown: true,
    },
  ]) {
    const request = vi.fn(async () => response)
    await expect(
      readPresentationResearchCapabilities(request, { includeRecovery: true }),
    ).rejects.toThrow('presentation_response_invalid')
    expect(request).toHaveBeenCalledTimes(1)
  }
  const network = vi.fn().mockRejectedValue(new Error('presentation_service_unavailable'))
  await expect(
    readPresentationResearchCapabilities(network, { includeRecovery: true }),
  ).rejects.toThrow('presentation_service_unavailable')
  expect(network).toHaveBeenCalledTimes(1)
  const missing = vi.fn().mockRejectedValue(new Error('presentation_upgrade_required'))
  await expect(
    readPresentationResearchCapabilities(missing, { includeRecovery: true }),
  ).rejects.toThrow('presentation_upgrade_required')
  expect(missing).toHaveBeenCalledTimes(3)
})
it('falls back from unsupported recovery to supported cleanup without legacy request', async () => {
  const request = vi
    .fn()
    .mockRejectedValueOnce(new Error('presentation_upgrade_required'))
    .mockResolvedValueOnce({
      version: 1,
      available: true,
      cleanupAvailable: true,
      historyVersions: [1, 2],
    })
  expect(await readPresentationResearchCapabilities(request, { includeRecovery: true })).toEqual({
    available: true,
    cleanupAvailable: true,
    historyVersion: 2,
  })
  expect(request).toHaveBeenCalledTimes(2)
})
