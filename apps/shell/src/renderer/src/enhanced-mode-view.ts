import type { EnhancedModeStatus } from '../../shared/enhanced-mode-api'

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
  const chinese = language === 'zh' || language === 'zh-TW'
  const label = chinese ? '增强模式' : 'Enhanced mode'
  const restartDetail = chinese ? '重启 WisWork 后生效' : 'Restart WisWork to apply'
  if (status.lifecycleState === 'blocked_by_policy' || status.lifecycleState === 'failed_safe') {
    return {
      label,
      detail:
        status.lifecycleState === 'blocked_by_policy'
          ? chinese
            ? '当前策略未开放增强模式'
            : 'Enhanced mode is disabled by policy'
          : chinese
            ? '增强模式启动失败，请切换到标准模式或重启后重试'
            : 'Enhanced mode failed safely; switch to Standard or restart to retry',
      action: status.requestedAgentRuntime === 'enhanced' ? 'disable' : 'none',
      actionLabel:
        status.requestedAgentRuntime === 'enhanced'
          ? chinese
            ? '切换到标准模式'
            : 'Switch to Standard mode'
          : '',
    }
  }
  if (!status.supported || status.component === 'unsupported') {
    return {
      label,
      detail: chinese ? '此设备暂不支持' : 'Not supported on this device',
      action: 'none',
      actionLabel: '',
    }
  }
  if (status.component === 'missing') {
    return {
      label,
      detail:
        status.requestedAgentRuntime === 'enhanced'
          ? chinese
            ? '使用前需要安装'
            : 'Install required before use'
          : chinese
            ? '需要下载可选组件'
            : 'Optional download required',
      action: 'install',
      actionLabel: chinese ? '下载' : 'Download',
    }
  }
  if (status.component === 'invalid') {
    return {
      label,
      detail: chinese ? '验证失败' : 'Verification failed',
      action: 'install',
      actionLabel: chinese ? '重新下载' : 'Download again',
    }
  }
  if (status.requestedAgentRuntime === 'enhanced') {
    return {
      label,
      detail: status.restartRequired ? restartDetail : chinese ? '增强模式' : 'Enhanced mode',
      action: 'disable',
      actionLabel: chinese ? '切换到标准模式' : 'Switch to Standard mode',
    }
  }
  return {
    label,
    detail: status.restartRequired ? restartDetail : chinese ? '标准模式' : 'Standard mode',
    action: 'enable',
    actionLabel: chinese ? '启用（重启后生效）' : 'Enable after restart',
    ...(status.restartRequired
      ? {}
      : {
          secondaryAction: 'remove' as const,
          secondaryActionLabel: chinese ? '移除可选组件' : 'Remove optional component',
        }),
  }
}
