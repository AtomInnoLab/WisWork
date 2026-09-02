import type { EnhancedModeApi, EnhancedModeStatus } from '../../shared/enhanced-mode-api'
import {
  normalizeLang,
  translateEnhancedMode,
  translatePresentationVerification,
} from '@wiswork/i18n'

export type EnhancedModeViewAction = 'none' | 'install' | 'enable' | 'disable'

export interface EnhancedModeView {
  readonly label: string
  readonly detail: string
  readonly standardLabel: string
  readonly enhancedLabel: string
  readonly selectedMode: 'standard' | 'enhanced'
  readonly activeMode: 'standard' | 'enhanced'
  readonly action: EnhancedModeViewAction
  readonly actionLabel: string
  readonly secondaryAction?: 'remove'
  readonly secondaryActionLabel?: string
}

export async function selectEnhancedMode(
  api: EnhancedModeApi,
  status: EnhancedModeStatus,
  target: 'standard' | 'enhanced',
): Promise<EnhancedModeStatus> {
  if (status.requestedAgentRuntime === target) return status
  if (target === 'standard') return api.setMode('standard')
  if (status.component === 'missing' || status.component === 'invalid') await api.install()
  return api.setMode('enhanced')
}

export function enhancedModeView(status: EnhancedModeStatus, language: string): EnhancedModeView {
  const lang = normalizeLang(language)
  const copy = (key: Parameters<typeof translateEnhancedMode>[1]) =>
    translateEnhancedMode(lang, key)
  const label = copy('label')
  const mode = {
    standardLabel: copy('standard'),
    enhancedLabel: label,
    selectedMode: status.requestedAgentRuntime,
    activeMode: status.activeAgentRuntime,
  } as const
  const restartDetail = copy('restart_required')
  if (status.lifecycleState === 'blocked_by_policy' || status.lifecycleState === 'failed_safe') {
    return {
      label,
      ...mode,
      detail:
        status.lifecycleState === 'blocked_by_policy'
          ? copy('blocked_by_policy')
          : copy('failed_safe'),
      action: status.requestedAgentRuntime === 'enhanced' ? 'disable' : 'none',
      actionLabel: status.requestedAgentRuntime === 'enhanced' ? copy('switch_standard') : '',
    }
  }
  if (!status.supported || status.component === 'unsupported') {
    return {
      label,
      ...mode,
      detail: copy('unavailable'),
      action: 'none',
      actionLabel: '',
    }
  }
  if (status.component === 'missing') {
    return {
      label,
      ...mode,
      detail:
        status.requestedAgentRuntime === 'enhanced'
          ? copy('install_required')
          : copy('optional_download'),
      action: 'install',
      actionLabel: copy('download'),
    }
  }
  if (status.component === 'invalid') {
    return {
      label,
      ...mode,
      detail:
        lang === 'en'
          ? 'Verification failed'
          : lang === 'zh' || lang === 'zh-TW'
            ? '验证失败'
            : translatePresentationVerification(lang, 'failed'),
      action: 'install',
      actionLabel: copy('download_again'),
    }
  }
  if (status.requestedAgentRuntime === 'enhanced') {
    return {
      label,
      ...mode,
      detail: status.restartRequired ? restartDetail : copy('enhanced'),
      action: 'disable',
      actionLabel: copy('switch_standard'),
    }
  }
  return {
    label,
    ...mode,
    detail: status.restartRequired ? restartDetail : copy('standard'),
    action: 'enable',
    actionLabel: copy('enable_after_restart'),
    ...(status.restartRequired
      ? {}
      : {
          secondaryAction: 'remove' as const,
          secondaryActionLabel: copy('remove'),
        }),
  }
}
