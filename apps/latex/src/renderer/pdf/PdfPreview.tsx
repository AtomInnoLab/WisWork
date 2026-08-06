import { useMemo } from 'react'
import { ReadonlyPdfViewer, type PdfPoint, type ViewerLocation } from '@wiswork/pdf-viewer'
import { useLatexLocale } from '../i18n/locale.js'

export interface PdfPreviewProps {
  pdfUrl: string | null
  revision: number | null
  location: ViewerLocation | null
  onReverseSync: (point: PdfPoint) => void
  stale: boolean
}

export function PdfPreview({ pdfUrl, revision, location, onReverseSync, stale }: PdfPreviewProps) {
  const { t } = useLatexLocale()
  const source = useMemo(() => (pdfUrl ? ({ kind: 'url', url: pdfUrl } as const) : null), [pdfUrl])
  return (
    <div className="pdf-preview">
      {stale && (
        <div className="preview-stale" role="status">
          {t('previewStale')}
        </div>
      )}
      <ReadonlyPdfViewer
        source={source}
        sourceKey={revision ?? 'empty'}
        location={location}
        onPageDoubleClick={onReverseSync}
      />
    </div>
  )
}
