import {
  presentationDeliveryBundleFiles,
  type PresentationDeliveryBundleReceipt,
} from '@wiswork/project-store/presentation-delivery-bundle'
export function hostBundleReceipt(
  documentId = 'document-1',
  projectId = 'project-1',
  requestId = 'pages',
): PresentationDeliveryBundleReceipt {
  const createdAt = '2026-09-29T00:00:00.000Z'
  return {
    version: 1,
    documentId,
    projectId,
    requestId,
    bundleId: 'a'.repeat(64),
    sha256: 'a'.repeat(64),
    sizeBytes: 100,
    receivedBytes: 100,
    state: 'ready',
    createdAt,
    completedAt: createdAt,
    manifest: {
      version: 1,
      scope: 'current_office_document',
      documentId,
      projectId,
      requestId,
      planRevision: 1,
      inputDigest: 'b'.repeat(64),
      planDigest: 'c'.repeat(64),
      createdAt,
      files: presentationDeliveryBundleFiles.map((name) => ({
        name,
        sizeBytes: 1,
        sha256: 'd'.repeat(64),
      })),
      checks: {
        completion: 'not_verified',
        sourceAuthority: 'not_verified',
        timeliness: 'not_verified',
        roundTrip: 'not_run',
        hostQa: 'historical_records_only',
        pdf: 'not_requested',
      },
    },
  }
}
