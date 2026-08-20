export type OfficeHost = 'word' | 'excel' | 'powerpoint' | 'unknown'

export interface OfficeAsyncResult {
  status: 'succeeded' | 'failed'
  value?: string
  message?: string
}

export interface OfficeRuntime {
  ready(): Promise<{ host?: unknown }>
  context: {
    document: {
      getSelectedDataAsync(
        coercionType: 'text',
        callback: (result: OfficeAsyncResult) => void,
      ): void
      setSelectedDataAsync(
        value: string,
        options: { coercionType: 'text' },
        callback: (result: OfficeAsyncResult) => void,
      ): void
    }
  }
}

export interface OfficeDocumentClient {
  initialize(): Promise<OfficeHost>
  readSelection(): Promise<string>
  replaceSelection(value: string): Promise<void>
  appendText(selection: string, value: string): Promise<void>
}

export function normalizeOfficeHost(host: unknown): OfficeHost {
  switch (typeof host === 'string' ? host.toLowerCase() : undefined) {
    case 'word':
      return 'word'
    case 'excel':
      return 'excel'
    case 'powerpoint':
      return 'powerpoint'
    default:
      return 'unknown'
  }
}

function officeError(result: OfficeAsyncResult): Error {
  return new Error(result.message || 'Office operation failed')
}

export function createOfficeDocumentClient(runtime: OfficeRuntime): OfficeDocumentClient {
  return {
    async initialize() {
      const info = await runtime.ready()
      return normalizeOfficeHost(info.host)
    },

    readSelection() {
      return new Promise((resolve, reject) => {
        runtime.context.document.getSelectedDataAsync('text', (result) => {
          if (result.status === 'failed') {
            reject(officeError(result))
            return
          }
          resolve(result.value ?? '')
        })
      })
    },

    replaceSelection(value) {
      return new Promise((resolve, reject) => {
        runtime.context.document.setSelectedDataAsync(value, { coercionType: 'text' }, (result) => {
          if (result.status === 'failed') {
            reject(officeError(result))
            return
          }
          resolve()
        })
      })
    },

    appendText(selection, value) {
      return new Promise((resolve, reject) => {
        runtime.context.document.setSelectedDataAsync(
          `${selection}${value}`,
          { coercionType: 'text' },
          (result) => {
            if (result.status === 'failed') {
              reject(officeError(result))
              return
            }
            resolve()
          },
        )
      })
    },
  }
}

export function createBrowserOfficeRuntime(): OfficeRuntime {
  return {
    async ready() {
      const info = await Office.onReady()
      return { host: info.host }
    },
    context: {
      document: {
        getSelectedDataAsync(_coercionType, callback) {
          Office.context.document.getSelectedDataAsync(
            Office.CoercionType.Text,
            (result: Office.AsyncResult<string>) =>
              callback({
                status:
                  result.status === Office.AsyncResultStatus.Succeeded ? 'succeeded' : 'failed',
                value: typeof result.value === 'string' ? result.value : '',
                message: result.error?.message,
              }),
          )
        },
        setSelectedDataAsync(value, _options, callback) {
          Office.context.document.setSelectedDataAsync(
            value,
            { coercionType: Office.CoercionType.Text },
            (result: Office.AsyncResult<void>) =>
              callback({
                status:
                  result.status === Office.AsyncResultStatus.Succeeded ? 'succeeded' : 'failed',
                message: result.error?.message,
              }),
          )
        },
      },
    },
  }
}
