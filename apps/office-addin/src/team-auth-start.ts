import { officeTeamAuthConfig } from './agent/team-auth-config.js'
import {
  validateTeamAuthorizationUrl,
  type OfficeTeamLoginDialogOptions,
} from './agent/team-login-dialog.js'
type Config = Pick<
  OfficeTeamLoginDialogOptions,
  'clientId' | 'authorizationEndpoint' | 'redirectUri'
>
export function startTeamAuthPage(
  location: { origin: string; hash: string; replace(value: string): void },
  config: Config | undefined,
): void {
  if (
    !config ||
    new URL(config.redirectUri).origin !== location.origin ||
    location.hash.length > 12_289
  )
    throw Error('team_auth_dialog_invalid')
  let authorizationUrl: string
  try {
    authorizationUrl = decodeURIComponent(location.hash.slice(1))
  } catch {
    throw Error('team_auth_dialog_invalid')
  }
  validateTeamAuthorizationUrl(authorizationUrl, config)
  location.replace(authorizationUrl)
}
if (typeof window !== 'undefined') {
  try {
    startTeamAuthPage(
      window.location,
      officeTeamAuthConfig(import.meta.env, window.location.origin),
    )
  } catch {
    const status = document.getElementById('team-auth-status')
    if (status) status.textContent = '团队登录无法开始，请关闭窗口并检查登录配置。'
  }
}
