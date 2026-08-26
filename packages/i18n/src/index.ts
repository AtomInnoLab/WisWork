export type Lang =
  | 'zh'
  | 'en'
  | 'ja'
  | 'ko'
  | 'fr'
  | 'de'
  | 'es'
  | 'th'
  | 'id'
  | 'ru'
  | 'ar'
  | 'pt'
  | 'it'
  | 'pl'
  | 'nl'
  | 'ms'
  | 'he'
  | 'hi'
  | 'zh-TW'

export const LANGS: readonly Lang[] = [
  'zh',
  'en',
  'ja',
  'ko',
  'fr',
  'de',
  'es',
  'th',
  'id',
  'ru',
  'ar',
  'pt',
  'it',
  'pl',
  'nl',
  'ms',
  'he',
  'hi',
  'zh-TW',
]

export function isLang(value: unknown): value is Lang {
  return typeof value === 'string' && (LANGS as readonly string[]).includes(value)
}

/** map a raw locale string ('zh-CN', 'zh-Hans', 'ja-JP', 'ko-KR', …) to a supported Lang */
export function normalizeLang(raw: string | null | undefined): Lang {
  const value = raw?.trim().toLowerCase()
  if (!value) return 'en'
  // traditional-script Chinese variants must win over the generic 'zh' prefix
  if (/^zh[-_](tw|hk|mo|hant)/.test(value)) return 'zh-TW'
  for (const lang of LANGS) {
    if (lang !== 'en' && lang !== 'zh-TW' && value.startsWith(lang)) return lang
  }
  // 'in' is the legacy ISO code for Indonesian still reported by some systems
  if (/^in\b/.test(value) || /^in[-_]/.test(value)) return 'id'
  // 'iw' is the legacy ISO code for Hebrew
  if (/^iw\b/.test(value) || /^iw[-_]/.test(value)) return 'he'
  return 'en'
}

const HTML_LANGS: Record<Lang, string> = {
  zh: 'zh-CN',
  en: 'en-US',
  ja: 'ja-JP',
  ko: 'ko-KR',
  fr: 'fr-FR',
  de: 'de-DE',
  es: 'es-ES',
  th: 'th-TH',
  id: 'id-ID',
  ru: 'ru-RU',
  ar: 'ar-SA',
  pt: 'pt-BR',
  it: 'it-IT',
  pl: 'pl-PL',
  nl: 'nl-NL',
  ms: 'ms-MY',
  he: 'he-IL',
  hi: 'hi-IN',
  'zh-TW': 'zh-TW',
}

/** BCP-47 tag for document.documentElement.lang (drives CSS :lang() and Chromium's per-language font fallback) */
export function htmlLang(lang: Lang): string {
  return HTML_LANGS[lang]
}

// ---- platform-native shortcut hints ----
// Dictionaries write shortcut hints in Mac notation (⌘S, ⇧⌘Z, ⌘+Click); on
// Windows/Linux every translated string is rewritten to Ctrl/Alt/Shift form.

const MAC_KEY_NAMES: Record<string, string> = {
  '⌫': 'Backspace',
  '⌦': 'Delete',
  '⏎': 'Enter',
  '↩': 'Enter',
  '␣': 'Space',
}

const HAS_MAC_SYMBOL = /[⌘⌃⌥⇧⌫⌦⏎↩␣]/
const CHORD = /([⌘⌃⌥⇧]+)(F\d{1,2}|[A-Za-z0-9±=`'\\,./;[\]\-←↑→↓⌫⌦⏎↩␣]|\+)?/g

function chordToWin(mods: string, key: string | undefined): string {
  const parts: string[] = []
  if (mods.includes('⌘') || mods.includes('⌃')) parts.push('Ctrl')
  if (mods.includes('⌥')) parts.push('Alt')
  if (mods.includes('⇧')) parts.push('Shift')
  if (key) parts.push(MAC_KEY_NAMES[key] ?? key)
  return parts.join('+')
}

/** rewrite Mac shortcut notation in a UI string to Windows/Linux form (pure) */
export function macShortcutsToWin(text: string): string {
  if (!HAS_MAC_SYMBOL.test(text)) return text
  return text
    .replace(/⌘\/(?=\p{L}{2})/gu, '') // "⌘/Ctrl+Enter" dual-platform listings: keep the Ctrl side
    .replace(CHORD, (_m, mods: string, key: string | undefined) =>
      key === '+' ? `${chordToWin(mods, undefined)}+` : chordToWin(mods, key),
    )
    .replace(/[⌫⌦⏎↩␣]/g, (glyph) => MAC_KEY_NAMES[glyph] ?? glyph)
}

const IS_MAC = (() => {
  const g = globalThis as {
    navigator?: { platform?: string }
    process?: { platform?: string }
  }
  if (g.navigator?.platform) return /mac/i.test(g.navigator.platform)
  return g.process?.platform === 'darwin'
})()

/** platform-aware shortcut display: identity on macOS */
export const platformShortcuts: (text: string) => string = IS_MAC
  ? (text) => text
  : macShortcutsToWin

export type Params = Record<string, string | number>

/** fill {name} placeholders; unknown placeholders are left as-is */
export function format(template: string, params?: Params): string {
  if (!params) return template
  return template.replace(/\{(\w+)\}/g, (match, name: string) =>
    name in params ? String(params[name]) : match,
  )
}

/** per-language dictionaries; zh defines the key set, all others must match it */
export type LangDicts<D extends Record<string, string>> = { zh: D } & {
  [L in Exclude<Lang, 'zh'>]: Record<keyof D, string>
}

/**
 * Identity helper for dictionary shards: keeps literal key inference while
 * type-checking that every other language covers exactly the zh key set.
 */
export function defineStrings<D extends Record<string, string>>(dicts: LangDicts<D>): LangDicts<D> {
  return dicts
}

// ---- process-wide current language ----
// Used by Electron main-process code (shell + editor main modules share one
// bundle, so one holder). Renderers get the language over IPC instead.

let uiLang: Lang = 'zh'
const langListeners = new Set<(lang: Lang) => void>()

export function getUiLang(): Lang {
  return uiLang
}

export function setUiLang(lang: Lang): void {
  if (lang === uiLang) return
  uiLang = lang
  for (const listener of langListeners) listener(lang)
}

export function onUiLangChange(listener: (lang: Lang) => void): () => void {
  langListeners.add(listener)
  return () => langListeners.delete(listener)
}

/**
 * Build a translator over per-language dictionaries. The zh dictionary defines
 * the key set; every other language must cover exactly the same keys
 * (compile-time checked), so a missing translation is a type error, not a
 * runtime fallback.
 */
export function createI18n<D extends Record<string, string>>(dicts: LangDicts<D>) {
  return (lang: Lang, key: keyof D, params?: Params): string =>
    platformShortcuts(format(dicts[lang][key], params))
}

export type ServiceErrorCode =
  | 'auth_required'
  | 'model_credentials_missing'
  | 'model_rate_limited'
  | 'model_upstream_unavailable'
  | 'model_invalid_response'

type ServiceErrorMessages = Record<ServiceErrorCode, string>

const serviceErrors: Record<Lang, ServiceErrorMessages> = {
  zh: {
    auth_required: '请登录 WisWork 后使用 AI。',
    model_credentials_missing: 'WisWork 模型服务尚未配置。',
    model_rate_limited: 'WisWork 模型服务繁忙，请稍后重试。',
    model_upstream_unavailable: 'WisWork 模型服务暂时不可用，请稍后重试。',
    model_invalid_response: 'WisWork 模型服务返回了无效响应，请稍后重试。',
  },
  en: {
    auth_required: 'Sign in to WisWork to use AI.',
    model_credentials_missing: 'The WisWork model service is not configured.',
    model_rate_limited: 'The WisWork model service is busy. Try again shortly.',
    model_upstream_unavailable: 'The WisWork model service is temporarily unavailable.',
    model_invalid_response: 'The WisWork model service returned an invalid response.',
  },
  ja: {
    auth_required: 'AI を使用するには WisWork にサインインしてください。',
    model_credentials_missing: 'WisWork モデルサービスが設定されていません。',
    model_rate_limited:
      'WisWork モデルサービスが混み合っています。しばらくしてから再試行してください。',
    model_upstream_unavailable: 'WisWork モデルサービスは一時的に利用できません。',
    model_invalid_response: 'WisWork モデルサービスから無効な応答が返されました。',
  },
  ko: {
    auth_required: 'AI를 사용하려면 WisWork에 로그인하세요.',
    model_credentials_missing: 'WisWork 모델 서비스가 구성되지 않았습니다.',
    model_rate_limited: 'WisWork 모델 서비스가 혼잡합니다. 잠시 후 다시 시도하세요.',
    model_upstream_unavailable: 'WisWork 모델 서비스를 일시적으로 사용할 수 없습니다.',
    model_invalid_response: 'WisWork 모델 서비스가 잘못된 응답을 반환했습니다.',
  },
  fr: {
    auth_required: 'Connectez-vous à WisWork pour utiliser l’IA.',
    model_credentials_missing: 'Le service de modèles WisWork n’est pas configuré.',
    model_rate_limited: 'Le service de modèles WisWork est occupé. Réessayez bientôt.',
    model_upstream_unavailable: 'Le service de modèles WisWork est temporairement indisponible.',
    model_invalid_response: 'Le service de modèles WisWork a renvoyé une réponse invalide.',
  },
  de: {
    auth_required: 'Melden Sie sich bei WisWork an, um KI zu verwenden.',
    model_credentials_missing: 'Der WisWork-Modelldienst ist nicht konfiguriert.',
    model_rate_limited: 'Der WisWork-Modelldienst ist ausgelastet. Versuchen Sie es später erneut.',
    model_upstream_unavailable: 'Der WisWork-Modelldienst ist vorübergehend nicht verfügbar.',
    model_invalid_response: 'Der WisWork-Modelldienst hat eine ungültige Antwort geliefert.',
  },
  es: {
    auth_required: 'Inicia sesión en WisWork para usar la IA.',
    model_credentials_missing: 'El servicio de modelos de WisWork no está configurado.',
    model_rate_limited:
      'El servicio de modelos de WisWork está ocupado. Inténtalo de nuevo pronto.',
    model_upstream_unavailable:
      'El servicio de modelos de WisWork no está disponible temporalmente.',
    model_invalid_response: 'El servicio de modelos de WisWork devolvió una respuesta no válida.',
  },
  th: {
    auth_required: 'ลงชื่อเข้าใช้ WisWork เพื่อใช้ AI',
    model_credentials_missing: 'ยังไม่ได้กำหนดค่าบริการโมเดล WisWork',
    model_rate_limited: 'บริการโมเดล WisWork ไม่ว่าง โปรดลองอีกครั้งภายหลัง',
    model_upstream_unavailable: 'บริการโมเดล WisWork ไม่พร้อมใช้งานชั่วคราว',
    model_invalid_response: 'บริการโมเดล WisWork ส่งการตอบกลับที่ไม่ถูกต้อง',
  },
  id: {
    auth_required: 'Masuk ke WisWork untuk menggunakan AI.',
    model_credentials_missing: 'Layanan model WisWork belum dikonfigurasi.',
    model_rate_limited: 'Layanan model WisWork sedang sibuk. Coba lagi nanti.',
    model_upstream_unavailable: 'Layanan model WisWork sementara tidak tersedia.',
    model_invalid_response: 'Layanan model WisWork memberikan respons yang tidak valid.',
  },
  ru: {
    auth_required: 'Войдите в WisWork, чтобы использовать ИИ.',
    model_credentials_missing: 'Служба моделей WisWork не настроена.',
    model_rate_limited: 'Служба моделей WisWork занята. Повторите попытку позже.',
    model_upstream_unavailable: 'Служба моделей WisWork временно недоступна.',
    model_invalid_response: 'Служба моделей WisWork вернула недопустимый ответ.',
  },
  ar: {
    auth_required: 'سجّل الدخول إلى WisWork لاستخدام الذكاء الاصطناعي.',
    model_credentials_missing: 'لم تتم تهيئة خدمة نماذج WisWork.',
    model_rate_limited: 'خدمة نماذج WisWork مشغولة. حاول مجددًا بعد قليل.',
    model_upstream_unavailable: 'خدمة نماذج WisWork غير متاحة مؤقتًا.',
    model_invalid_response: 'أعادت خدمة نماذج WisWork استجابة غير صالحة.',
  },
  pt: {
    auth_required: 'Entre no WisWork para usar a IA.',
    model_credentials_missing: 'O serviço de modelos do WisWork não está configurado.',
    model_rate_limited: 'O serviço de modelos do WisWork está ocupado. Tente novamente em breve.',
    model_upstream_unavailable:
      'O serviço de modelos do WisWork está temporariamente indisponível.',
    model_invalid_response: 'O serviço de modelos do WisWork retornou uma resposta inválida.',
  },
  it: {
    auth_required: 'Accedi a WisWork per usare l’IA.',
    model_credentials_missing: 'Il servizio modelli WisWork non è configurato.',
    model_rate_limited: 'Il servizio modelli WisWork è occupato. Riprova tra poco.',
    model_upstream_unavailable: 'Il servizio modelli WisWork è temporaneamente non disponibile.',
    model_invalid_response: 'Il servizio modelli WisWork ha restituito una risposta non valida.',
  },
  pl: {
    auth_required: 'Zaloguj się do WisWork, aby korzystać z AI.',
    model_credentials_missing: 'Usługa modeli WisWork nie jest skonfigurowana.',
    model_rate_limited: 'Usługa modeli WisWork jest zajęta. Spróbuj ponownie później.',
    model_upstream_unavailable: 'Usługa modeli WisWork jest tymczasowo niedostępna.',
    model_invalid_response: 'Usługa modeli WisWork zwróciła nieprawidłową odpowiedź.',
  },
  nl: {
    auth_required: 'Meld u aan bij WisWork om AI te gebruiken.',
    model_credentials_missing: 'De WisWork-modelservice is niet geconfigureerd.',
    model_rate_limited: 'De WisWork-modelservice is bezet. Probeer het later opnieuw.',
    model_upstream_unavailable: 'De WisWork-modelservice is tijdelijk niet beschikbaar.',
    model_invalid_response: 'De WisWork-modelservice heeft een ongeldig antwoord gegeven.',
  },
  ms: {
    auth_required: 'Log masuk ke WisWork untuk menggunakan AI.',
    model_credentials_missing: 'Perkhidmatan model WisWork belum dikonfigurasi.',
    model_rate_limited: 'Perkhidmatan model WisWork sedang sibuk. Cuba lagi sebentar lagi.',
    model_upstream_unavailable: 'Perkhidmatan model WisWork tidak tersedia buat sementara waktu.',
    model_invalid_response: 'Perkhidmatan model WisWork mengembalikan respons yang tidak sah.',
  },
  he: {
    auth_required: 'יש להתחבר ל-WisWork כדי להשתמש ב-AI.',
    model_credentials_missing: 'שירות המודלים של WisWork אינו מוגדר.',
    model_rate_limited: 'שירות המודלים של WisWork עמוס. נסו שוב מאוחר יותר.',
    model_upstream_unavailable: 'שירות המודלים של WisWork אינו זמין זמנית.',
    model_invalid_response: 'שירות המודלים של WisWork החזיר תגובה לא תקינה.',
  },
  hi: {
    auth_required: 'AI का उपयोग करने के लिए WisWork में साइन इन करें।',
    model_credentials_missing: 'WisWork मॉडल सेवा कॉन्फ़िगर नहीं है।',
    model_rate_limited: 'WisWork मॉडल सेवा व्यस्त है। थोड़ी देर बाद फिर प्रयास करें।',
    model_upstream_unavailable: 'WisWork मॉडल सेवा अस्थायी रूप से उपलब्ध नहीं है।',
    model_invalid_response: 'WisWork मॉडल सेवा ने अमान्य प्रतिक्रिया लौटाई।',
  },
  'zh-TW': {
    auth_required: '請登入 WisWork 後使用 AI。',
    model_credentials_missing: 'WisWork 模型服務尚未設定。',
    model_rate_limited: 'WisWork 模型服務忙碌中，請稍後再試。',
    model_upstream_unavailable: 'WisWork 模型服務暫時無法使用。',
    model_invalid_response: 'WisWork 模型服務傳回無效回應。',
  },
}

export function translateServiceError(lang: Lang, code: ServiceErrorCode): string {
  return serviceErrors[lang][code]
}
