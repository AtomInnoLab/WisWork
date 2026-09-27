import { createHash } from 'node:crypto'
import type { PresentationInlineAsset } from '@wiswork/pptx-engine/presentation'
import type { PresentationPlan } from '@wiswork/pptx-engine/presentation-plan'

/** Bind the brand logo to the bytes actually passed to the compiler, after attachment conversion. */
export function assertBrandLogoAsset(plan: PresentationPlan, assets: PresentationInlineAsset[]): void {
  const logo = plan.brandKit?.logo
  if (!logo) return
  const asset = assets.find((item) => item.id === logo.assetId)
  if (!asset || !/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(asset.base64))
    throw new Error('plan_mismatch')
  const bytes = Buffer.from(asset.base64, 'base64')
  if (!bytes.length || createHash('sha256').update(bytes).digest('hex') !== logo.assetDigest)
    throw new Error('plan_mismatch')
}
