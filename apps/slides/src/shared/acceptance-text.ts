import { fingerprintSemanticValue } from '@wiswork/presentation-ops'

const normalizeText = (text: string): string =>
  text.replaceAll('\r\n', '\n').replaceAll('\r', '\n').normalize('NFC')

export async function digestSlidesAcceptanceText(
  leaseToken: string,
  checkId: string,
  targetToken: string,
  text: string,
): Promise<string> {
  if (text.length > 1_000_000) throw new TypeError('Acceptance text is overbound')
  return fingerprintSemanticValue({ leaseToken, checkId, targetToken, text: normalizeText(text) })
}
