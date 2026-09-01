import { presentationVerificationFlags } from '@wiswork/presentation-verification'

type PresentationFlagEnv = Record<string, string | undefined>

export const officePresentationVerificationFlags = (env: PresentationFlagEnv) =>
  presentationVerificationFlags(env, 'VITE_WISWORK_PRESENTATION_')
