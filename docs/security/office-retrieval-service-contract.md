# Office retrieval service activation contract

Office web tools remain disabled unless the PC build contains an exact canonical HTTPS endpoint in
`OFFICE_RETRIEVAL_SERVICES`. An environment variable can select a compiled entry but cannot add a
destination. Each entry attests both `wiswork-office-retrieval-v1` response schemas and
`dns-rebinding-and-redirect-hops-v1` SSRF protection. The production map is intentionally empty
until the service owner publishes and reviews that contract.

The PC rejects malformed requests and recognized non-public IP literals as defense in depth. It
does **not** claim to enforce DNS-level SSRF protection. The fixed retrieval service must:

- resolve every hostname itself and reject every non-public answer before connecting;
- pin or revalidate the selected address so DNS rebinding cannot change the connection target;
- disable automatic redirects, then repeat URL, scheme, DNS, and address validation for every hop;
- allow HTTPS only and reject credentials, local/link-local/private/reserved destinations, unsafe
  ports, unsupported content types, oversized bodies, excess redirects, and deadline overruns;
- authenticate the WisWork PC identity without returning credentials, upstream bodies, or raw
  errors to Office.

Deployment order is Relay v2, PC with explicit `pc.negotiate`, the attested retrieval service, then
Office v2 capability enablement. Older Office v1 pairings remain usable because PC obtains the
pairing version from Relay before sending `pc.claim`; it never guesses or silently downgrades.
