interface PrivilegedSchemeProtocol {
  registerSchemesAsPrivileged(
    schemes: Array<{
      scheme: string
      privileges: { secure: boolean; standard: boolean; supportFetchAPI: boolean }
    }>,
  ): void
}

const registeredOwners = new WeakSet<object>()

export function registerLatexProtocolScheme(protocol: PrivilegedSchemeProtocol): void {
  if (registeredOwners.has(protocol)) return
  protocol.registerSchemesAsPrivileged([
    {
      scheme: 'wiswork-latex-pdf',
      privileges: { secure: true, standard: true, supportFetchAPI: true },
    },
  ])
  registeredOwners.add(protocol)
}
