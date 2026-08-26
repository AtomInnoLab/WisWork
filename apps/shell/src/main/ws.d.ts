declare module 'ws' {
  export default class WebSocket {
    constructor(url: string, options: { headers: Record<string, string> })
    readyState: number
    addEventListener(name: string, listener: (event: unknown) => void): void
    send(data: string): void
    close(code?: number, reason?: string): void
  }
}
