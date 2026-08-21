export const PREFERRED_OFFICE_BRIDGE_PORT = 43_127
export const OFFICE_RELAY_CONNECT_ORIGIN = 'wss://office.8-216-134-194.sslip.io'
export const DEFAULT_OFFICE_BRIDGE_PORTS = Object.freeze([
  PREFERRED_OFFICE_BRIDGE_PORT,
  ...Array.from({ length: 64 }, (_, index) => 43_120 + index).filter(
    (port) => port !== PREFERRED_OFFICE_BRIDGE_PORT,
  ),
])

type BuildEnv = Record<string, string | undefined>
export interface DeploymentConfig {
  addinOrigin: string
  bridgePorts: readonly number[]
}

export type OfficeTransportMode = 'relay' | 'loopback'
export type OfficeWorkspaceMode = 'workspace' | 'legacy'
export interface OfficeCapabilityFlags {
  conversions: boolean
  skillPackages: boolean
  importMedia: boolean
  webTools: boolean
}

function exactFlag(value: string | undefined, defaultValue: boolean): boolean {
  if (value === undefined || value === '') return defaultValue
  if (value === '1') return true
  if (value === '0') return false
  throw new Error('invalid_office_capability_flags')
}

export function officeCapabilityFlags(env: BuildEnv): OfficeCapabilityFlags {
  return Object.freeze({
    conversions: exactFlag(env.VITE_WISWORK_OFFICE_CONVERSIONS, true),
    skillPackages: exactFlag(env.VITE_WISWORK_OFFICE_SKILL_PACKAGES, true),
    importMedia: exactFlag(env.VITE_WISWORK_OFFICE_IMPORT_MEDIA, true),
    webTools: exactFlag(env.VITE_WISWORK_OFFICE_WEB_TOOLS, false),
  })
}

export function officeWorkspaceMode(env: BuildEnv): OfficeWorkspaceMode {
  const value = env.VITE_WISWORK_OFFICE_WORKSPACE
  if (value === undefined || value === '' || value === '1') return 'workspace'
  if (value === '0') return 'legacy'
  throw new Error('invalid_office_workspace_mode')
}

export function officeTransportMode(env: BuildEnv): OfficeTransportMode {
  const value = env.VITE_WISWORK_OFFICE_TRANSPORT
  if (value === undefined || value === '' || value === 'relay') return 'relay'
  if (value === 'loopback') return 'loopback'
  throw new Error('invalid_office_transport')
}

export function officeBridgePorts(env: BuildEnv): readonly number[] | undefined {
  const configured = env.VITE_WISWORK_PC_BRIDGE_PORTS
  if (configured === undefined) return DEFAULT_OFFICE_BRIDGE_PORTS
  if (configured !== configured.trim()) return undefined
  const values = configured.split(',')
  if (values.length < 1 || values.length > 128) return undefined
  if (values.some((value) => !/^(?:[1-9]\d{0,4})$/.test(value))) return undefined
  const ports = values.map(Number)
  if (ports.some((port) => port > 65_535) || new Set(ports).size !== ports.length) return undefined
  return Object.freeze(ports)
}

export function officeBridgeEndpoints(env: BuildEnv): readonly string[] {
  const ports = officeBridgePorts(env)
  if (!ports) throw new Error('invalid_office_bridge_ports')
  return ports.map((port) => `http://127.0.0.1:${port}`)
}

export function deploymentConfig(env: BuildEnv): DeploymentConfig | undefined {
  const value = env.VITE_WISWORK_ADDIN_ORIGIN?.trim()
  if (!value) return undefined
  let mode: OfficeTransportMode
  try {
    mode = officeTransportMode(env)
    void officeWorkspaceMode(env)
    void officeCapabilityFlags(env)
  } catch {
    return undefined
  }
  const bridgePorts = officeBridgePorts(env)
  if (!bridgePorts) return undefined
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
    void mode
    return { addinOrigin: url.origin, bridgePorts }
  } catch {
    return undefined
  }
}

export function deploymentConnectOrigins(env: BuildEnv): string {
  try {
    return officeTransportMode(env) === 'relay'
      ? OFFICE_RELAY_CONNECT_ORIGIN
      : officeBridgeEndpoints(env).join(' ')
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
