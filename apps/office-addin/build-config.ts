export const DEFAULT_OFFICE_BRIDGE_PORT = 43_127

type BuildEnv = Record<string, string | undefined>
export interface DeploymentConfig {
  addinOrigin: string
  bridgePort: number
}

export function officeBridgePort(env: BuildEnv): number | undefined {
  const raw = env.VITE_WISWORK_PC_BRIDGE_PORT?.trim()
  if (!raw) return DEFAULT_OFFICE_BRIDGE_PORT
  if (!/^\d+$/.test(raw)) return undefined
  const port = Number(raw)
  return Number.isSafeInteger(port) && port >= 1 && port <= 65_535 ? port : undefined
}

export function officeBridgeEndpoint(env: BuildEnv): string {
  const port = officeBridgePort(env)
  if (port === undefined) throw new Error('invalid_office_bridge_port')
  return `http://127.0.0.1:${port}`
}

export function deploymentConfig(env: BuildEnv): DeploymentConfig | undefined {
  const value = env.VITE_WISWORK_ADDIN_ORIGIN?.trim()
  if (!value) return undefined
  const bridgePort = officeBridgePort(env)
  if (bridgePort === undefined) return undefined
  try {
    const url = new URL(value)
    if (
      url.protocol !== 'https:' ||
      url.origin !== value ||
      url.username ||
      url.password ||
      url.hostname.includes('*')
    )
      return undefined
    return { addinOrigin: url.origin, bridgePort }
  } catch {
    return undefined
  }
}

export function deploymentConnectOrigins(env: BuildEnv): string {
  try {
    return officeBridgeEndpoint(env)
  } catch {
    return ''
  }
}

export function renderDeploymentManifest(template: string, config: DeploymentConfig): string {
  return template
    .replace(
      '<!-- DEVELOPMENT-ONLY MANIFEST: production builds generate origin-specific dist/manifest.xml. -->\n',
      '',
    )
    .replaceAll('https://localhost:3000', config.addinOrigin)
    .replace(
      / {2}<AppDomains>[\s\S]*?<\/AppDomains>/,
      `  <AppDomains>\n    <AppDomain>${config.addinOrigin}</AppDomain>\n  </AppDomains>`,
    )
}
