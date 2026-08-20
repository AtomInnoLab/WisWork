export const OFFICE_BRIDGE_ORIGIN = 'http://127.0.0.1:43127'

type BuildEnv = Record<string, string | undefined>
export interface DeploymentConfig {
  addinOrigin: string
}

export function deploymentConfig(env: BuildEnv): DeploymentConfig | undefined {
  const value = env.VITE_WISWORK_ADDIN_ORIGIN?.trim()
  if (!value) return undefined
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
    return { addinOrigin: url.origin }
  } catch {
    return undefined
  }
}

export function deploymentConnectOrigins(): string {
  return OFFICE_BRIDGE_ORIGIN
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
