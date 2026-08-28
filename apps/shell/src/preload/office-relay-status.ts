import type { OfficeRelayStatus } from '../shared/home-api'

const SAFE_OFFICE_RELAY_STATUSES = new Set<OfficeRelayStatus>([
  'disconnected',
  'connecting',
  'claiming',
  'awaiting_approval',
  'waiting_for_office',
  'paired',
  'disconnected:auth_required',
  'disconnected:logout',
  'disconnected:network_error',
  'disconnected:new_claim',
  'disconnected:new_resume',
  'disconnected:new_revocation',
  'disconnected:pairing_expired',
  'disconnected:protocol_violation',
  'disconnected:rejected',
  'disconnected:relay_error',
  'disconnected:relay_closed',
  'disconnected:session_expired',
  'disconnected:binding_unavailable',
  'disconnected:binding_revoked',
  'disconnected:binding_not_remembered',
  'disconnected:capability_not_negotiated',
  'disconnected:peer_unavailable',
  'disconnected:resume_limit',
  'disconnected:resume_rate_limited',
  'disconnected:account_switch',
  'disconnected:shutdown',
  'error:invalid_config',
  'error:binding_lifecycle',
])

export function sanitizeOfficeRelayStatus(value: unknown): OfficeRelayStatus {
  return typeof value === 'string' && SAFE_OFFICE_RELAY_STATUSES.has(value as OfficeRelayStatus)
    ? (value as OfficeRelayStatus)
    : 'disconnected'
}
