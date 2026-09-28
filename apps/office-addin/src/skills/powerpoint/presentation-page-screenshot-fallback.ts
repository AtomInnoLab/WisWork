import {
  validatePowerPointPageScreenshot,
  type BrowserPowerPointAdapter,
  type PowerPointPageInspection,
} from './browser-powerpoint-adapter.js'
import { describePagePackageBackup, saveChartPackageBackup } from './presentation-chart-backup.js'
import { presentationPackageDigest } from './powerpoint-package.js'

type Request = (body: unknown, signal?: AbortSignal) => Promise<Response>
type Options = {
  adapter: BrowserPowerPointAdapter
  request: Request
  documentId(): Promise<string>
}
const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b)
const screenshotFailure = (error: unknown) =>
  !!error &&
  typeof error === 'object' &&
  ['office_screenshot_unavailable', 'ActivityLimitReached', 'Timeout'].includes(
    String((error as { code?: unknown }).code),
  )

async function cleanupBackup(
  request: Request,
  input: {
    backupId: string
    documentId: string
    hostSlideId: string
    slideIds: string[]
    sha256: string
    sizeBytes: number
  },
): Promise<void> {
  for (let attempt = 0; attempt < 2; attempt++) {
    try {
      const statusResponse = await request({
        operation: 'existing_page_backup_status',
        documentId: input.documentId,
        backupId: input.backupId,
      })
      const status: unknown = await statusResponse.json()
      if (
        status &&
        typeof status === 'object' &&
        (status as Record<string, unknown>).error === 'not_found'
      )
        return
      if (
        !statusResponse.ok ||
        !status ||
        typeof status !== 'object' ||
        Array.isArray(status) ||
        (status as Record<string, unknown>).backupId !== input.backupId ||
        (status as Record<string, unknown>).documentId !== input.documentId ||
        (status as Record<string, unknown>).hostSlideId !== input.hostSlideId ||
        !same((status as Record<string, unknown>).slideIds, input.slideIds) ||
        (status as Record<string, unknown>).sha256 !== input.sha256 ||
        (status as Record<string, unknown>).sizeBytes !== input.sizeBytes
      )
        continue
      const operation =
        (status as Record<string, unknown>).status === 'ready'
          ? 'existing_page_backup_release'
          : (status as Record<string, unknown>).status === 'uploading'
            ? 'existing_page_backup_abandon'
            : undefined
      if (!operation) continue
      const response = await request({ operation, ...input })
      const receipt: unknown = await response.json()
      if (
        response.ok &&
        receipt &&
        typeof receipt === 'object' &&
        (receipt as Record<string, unknown>).backupId === input.backupId &&
        (receipt as Record<string, unknown>).documentId === input.documentId &&
        (receipt as Record<string, unknown>).status ===
          (operation === 'existing_page_backup_release' ? 'released' : 'abandoned')
      )
        return
    } catch {
      // Both cleanup operations are scoped and retryable.
    }
  }
  throw new Error('presentation_page_backup_cleanup_failed')
}

export function createPresentationPageScreenshotFallback(options: Options) {
  return async (slideId: string, signal?: AbortSignal): Promise<PowerPointPageInspection> => {
    try {
      return await options.adapter.inspectPresentationPage(slideId, signal)
    } catch (hostError) {
      if (!screenshotFailure(hostError) || signal?.aborted) throw hostError
      try {
        const documentId = await options.documentId()
        const exported = await options.adapter.exportPresentationPagePackage(slideId, signal)
        if (exported.slideId !== slideId)
          throw new Error('presentation_qa_stale', { cause: hostError })
        const backupId = `qa_${crypto.randomUUID().replaceAll('-', '')}`
        const described = await describePagePackageBackup(exported.base64, signal)
        const packageDigest = described.packageDigest
        const scope = {
          request: options.request,
          documentId,
          hostSlideId: slideId,
          slideIds: exported.slideIds,
        }
        const cleanup = () =>
          cleanupBackup(options.request, {
            backupId,
            documentId,
            hostSlideId: slideId,
            slideIds: exported.slideIds,
            sha256: described.sha256,
            sizeBytes: described.sizeBytes,
          })
        let backup: Awaited<ReturnType<typeof saveChartPackageBackup>>
        try {
          backup = await saveChartPackageBackup(
            { ...scope, base64: exported.base64, backupId },
            signal,
          )
        } catch (error) {
          await cleanup()
          throw error
        }
        let base64: string
        try {
          const response = await options.request(
            { operation: 'existing_page_backup_render', documentId, backupId },
            signal,
          )
          if (!response.ok) throw new Error('renderer_unavailable', { cause: hostError })
          const value: unknown = await response.json()
          if (
            !value ||
            typeof value !== 'object' ||
            Array.isArray(value) ||
            (value as Record<string, unknown>).backupId !== backupId ||
            (value as Record<string, unknown>).hostSlideId !== slideId ||
            (value as Record<string, unknown>).sha256 !== backup.sha256 ||
            (value as Record<string, unknown>).renderer !== 'libreoffice' ||
            (value as Record<string, unknown>).mime !== 'image/png'
          )
            throw new Error('renderer_unavailable', { cause: hostError })
          base64 = validatePowerPointPageScreenshot((value as Record<string, unknown>).base64)
          if (atob(base64).length > 64 * 1024)
            throw new Error('renderer_unavailable', { cause: hostError })
        } finally {
          // The render transport is a temporary read-only use of the existing backup store.
          await cleanup()
        }
        const assertCurrent = async () => {
          const current = await options.adapter.exportPresentationPagePackage(slideId, signal)
          if (
            (await options.documentId()) !== documentId ||
            current.slideId !== slideId ||
            !same(current.slideIds, exported.slideIds) ||
            (await presentationPackageDigest(current.base64, signal)) !== packageDigest
          )
            throw new Error('presentation_qa_stale')
        }
        await assertCurrent()
        const inspected = await options.adapter.inspectPresentationPage(slideId, signal, base64)
        await assertCurrent()
        return inspected
      } catch (fallbackError) {
        if (
          fallbackError instanceof Error &&
          [
            'cancelled',
            'office_concurrent_change',
            'presentation_qa_stale',
            'presentation_page_backup_cleanup_failed',
          ].includes(fallbackError.message)
        )
          throw fallbackError
        throw hostError
      }
    }
  }
}
