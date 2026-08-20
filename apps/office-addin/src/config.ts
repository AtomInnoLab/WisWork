export interface RuntimeConfig {
  authorizationUrl: string
  tokenUrl: string
  callbackUrl: string
  clientId: string
  issuer: string
  messagesUrl: string
}

export type RuntimeConfigState =
  { status: 'available'; config: RuntimeConfig } | { status: 'unavailable' }

type RuntimeEnv = Record<string, string | undefined>

const ENVIRONMENT_KEYS = {
  authorizationUrl: 'VITE_WISWORK_AUTHORIZATION_URL',
  tokenUrl: 'VITE_WISWORK_TOKEN_URL',
  callbackUrl: 'VITE_WISWORK_CALLBACK_URL',
  clientId: 'VITE_WISWORK_CLIENT_ID',
  issuer: 'VITE_WISWORK_ISSUER',
  messagesUrl: 'VITE_WISWORK_MESSAGES_URL',
} as const

function validUrl(value: string, requireHttps: boolean): boolean {
  try {
    const url = new URL(value)
    const permittedProtocol = requireHttps
      ? url.protocol === 'https:'
      : url.protocol === 'https:' || url.protocol === 'http:'
    return !url.username && !url.password && permittedProtocol
  } catch {
    return false
  }
}

export function loadRuntimeConfig(
  env: RuntimeEnv,
  options: { production: boolean },
): RuntimeConfigState {
  const values = Object.fromEntries(
    Object.entries(ENVIRONMENT_KEYS).map(([name, key]) => [name, env[key]?.trim() ?? '']),
  ) as unknown as RuntimeConfig

  if (
    !values.clientId ||
    !values.issuer ||
    !validUrl(values.authorizationUrl, true) ||
    !validUrl(values.tokenUrl, true) ||
    !validUrl(values.issuer, true) ||
    !validUrl(values.messagesUrl, true) ||
    !validUrl(values.callbackUrl, options.production)
  ) {
    return { status: 'unavailable' }
  }

  return { status: 'available', config: values }
}
