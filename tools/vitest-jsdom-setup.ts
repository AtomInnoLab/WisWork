/**
 * Node 25 exposes an experimental global localStorage accessor even when no
 * persistence file is configured. Vitest sees the existing property and does
 * not replace it with jsdom's implementation, leaving tests with an object
 * that has no Storage methods. Install the environment-owned implementation
 * explicitly so browser tests behave consistently across supported Node
 * versions.
 */
const environment = globalThis as typeof globalThis & {
  jsdom?: { window: Window }
}

if (environment.jsdom) {
  Object.defineProperty(globalThis, 'localStorage', {
    value: environment.jsdom.window.localStorage,
    configurable: true,
    writable: true,
  })
}
