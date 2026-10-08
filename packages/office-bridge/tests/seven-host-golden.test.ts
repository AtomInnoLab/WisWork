import { describe, expect, it } from 'vitest'
import { selectOfficeAgentRuntime, type OfficeEnhancedHost } from '../src/index'

const prove = (host: OfficeEnhancedHost) => {
  const statement = {
    version: 1 as const,
    runtime_mode: 'enhanced' as const,
    runtime_instance: `runtime_${host}_0123456789`,
    component_version: '0.147.0',
    host,
    raw_office: true,
    expires_at: 2_000,
    policy_generation: 7,
    session_generation: 3,
  }
  expect(
    selectOfficeAgentRuntime(statement, {
      host,
      now: 1_000,
      componentVersion: '0.147.0',
      sessionGeneration: 3,
      minimumPolicyGeneration: 7,
    }),
  ).toEqual({
    mode: 'enhanced',
    rawOfficeAllowed: true,
    runtimeInstance: statement.runtime_instance,
  })
}

describe('independent Office Enhanced bridge goldens', () => {
  it('office-word Enhanced bridge', () => prove('office-word'))
  it('office-excel Enhanced bridge', () => prove('office-excel'))
  it('office-powerpoint Enhanced bridge', () => prove('office-powerpoint'))
})
