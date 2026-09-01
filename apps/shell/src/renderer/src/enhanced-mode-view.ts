import type { EnhancedModeStatus } from '../../shared/enhanced-mode-api'
import {
  normalizeLang,
  translateEnhancedMode,
  translatePresentationVerification,
} from '@wiswork/i18n'

export type EnhancedModeViewAction = 'none' | 'install' | 'enable' | 'disable'

export interface EnhancedModeView {
  readonly label: string
  readonly detail: string
  readonly action: EnhancedModeViewAction
  readonly actionLabel: string
  readonly secondaryAction?: 'remove'
  readonly secondaryActionLabel?: string
}

export function enhancedModeView(status: EnhancedModeStatus, language: string): EnhancedModeView {
  const lang = normalizeLang(language)
  const copy = (key: Parameters<typeof translateEnhancedMode>[1]) =>
    translateEnhancedMode(lang, key)
  const label = copy('label')
  const restartDetail = copy('restart_required')
  if (status.lifecycleState === 'blocked_by_policy' || status.lifecycleState === 'failed_safe') {
    return {
      label,
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
      detail: copy('unavailable'),
      action: 'none',
      actionLabel: '',
    }
  }
  if (status.component === 'missing') {
    return {
      label,
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
      detail: status.restartRequired ? restartDetail : copy('enhanced'),
      action: 'disable',
      actionLabel: copy('switch_standard'),
    }
  }
  return {
    label,
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
