import {
  replacePowerPointShapeGeometryPackage,
  type NativePageGeometry,
} from './presentation-geometry-revision-package.js'
import { replacePowerPointPictureMediaPackage } from './presentation-picture-package.js'
import { replacePowerPointTextRangePackage } from './presentation-text-revision-package.js'

interface CompositeRevision {
  text: { shapeId: string; start: number; before: string; after: string }
  geometry: { shapeId: string; before: NativePageGeometry; after: NativePageGeometry }
  picture: { shapeId: string; image: { mime: 'image/png' | 'image/jpeg'; base64: string } }
}

/** Build one native single-page revision for three distinct targets before any host write. */
export async function preparePowerPointCompositePagePackage(
  source: string,
  revision: CompositeRevision,
  signal?: AbortSignal,
): Promise<{
  base64: string
  beforeDigest: string
  afterDigest: string
  changedRuns: number
  mediaDigest: string
}> {
  const ids = [revision.text.shapeId, revision.geometry.shapeId, revision.picture.shapeId]
  if (new Set(ids).size !== 3) throw new Error('invalid_tool_input')
  const text = await replacePowerPointTextRangePackage(
    source,
    revision.text.shapeId,
    revision.text.start,
    revision.text.before,
    revision.text.after,
    signal,
  )
  const geometry = await replacePowerPointShapeGeometryPackage(
    text.base64,
    revision.geometry.shapeId,
    revision.geometry.before,
    revision.geometry.after,
    signal,
  )
  const picture = await replacePowerPointPictureMediaPackage(
    geometry.base64,
    revision.picture.shapeId,
    revision.picture.image,
    signal,
  )
  return {
    base64: picture.base64,
    beforeDigest: text.beforeDigest,
    afterDigest: picture.afterDigest,
    changedRuns: text.changedRuns,
    mediaDigest: picture.mediaDigest,
  }
}
