import { presentationVerificationFlags } from '@wiswork/presentation-verification'

type PresentationFlagEnv = Record<string, string | undefined>

export const officePresentationVerificationFlags = (env: PresentationFlagEnv) => {
  const flags = presentationVerificationFlags(env, 'VITE_WISWORK_PRESENTATION_')
  return Object.freeze({
    ...flags,
    // Office PowerPoint now follows the desktop Slides quality loop by default. The exact
    // environment flag remains an emergency rollback switch.
    autoCorrection: !env.VITE_WISWORK_PRESENTATION_AUTO_CORRECTION ? true : flags.autoCorrection,
  })
}
