export function captureAndScrubOAuthCallback(
  expectedCallbackUrl: string,
  currentHref: string,
  replace: (cleanUrl: string) => void,
): string | undefined {
  const expected = new URL(expectedCallbackUrl)
  const current = new URL(currentHref)
  if (current.origin !== expected.origin || current.pathname !== expected.pathname) return undefined

  const captured = current.href
  replace(new URL('/taskpane.html', expected.origin).href)
  return captured
}
