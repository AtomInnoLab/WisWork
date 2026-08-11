export function LoginDiagnostic({ value, onCopy }: { value: string; onCopy: () => void }) {
  return (
    <section className="login-diagnostic" aria-label="Login diagnostic">
      <code className="login-diagnostic-value">{value}</code>
      <button type="button" onClick={onCopy}>
        复制诊断 / Copy
      </button>
    </section>
  )
}
