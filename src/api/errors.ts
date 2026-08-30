export type AppErrorCode =
  | 'NETWORK_ERROR'
  | 'CORS_BLOCKED'
  | 'UNAUTHORIZED'
  | 'RATE_LIMITED'
  | 'MODEL_NOT_FOUND'
  | 'MODEL_NO_TOOL_SUPPORT'
  | 'REQUEST_TIMEOUT'
  | 'AGENT_LOOP'
  | 'AGENT_BUDGET'
  | 'DOCUMENT_CONFLICT'
  | 'NON_TEXT_OBJECTS'
  | 'WORD_API_UNSUPPORTED'
  | 'REQUEST_FAILED'
  | 'USER_CANCELLED'

export class AppError extends Error {
  code: AppErrorCode
  status?: number
  endpoint?: string
  retryable: boolean
  details?: Record<string, string | number | boolean>

  constructor(
    code: AppErrorCode,
    message: string,
    options: {
      status?: number
      endpoint?: string
      retryable?: boolean
      details?: Record<string, string | number | boolean>
    } = {},
  ) {
    super(message)
    this.name = 'AppError'
    this.code = code
    this.status = options.status
    this.endpoint = options.endpoint
    this.retryable = options.retryable ?? false
    this.details = options.details
  }
}

const getErrorStatus = (error: unknown): number | undefined => {
  if (error instanceof AppError) return error.status
  if (typeof error === 'object' && error && 'status' in error && typeof error.status === 'number') return error.status
  return undefined
}

const getErrorMessage = (error: unknown): string => {
  if (error instanceof Error) return error.message
  return String(error)
}

export const classifyError = (error: unknown, endpoint = ''): AppError => {
  if (error instanceof AppError) return error

  const message = getErrorMessage(error)
  const normalized = message.toLowerCase()
  const status = getErrorStatus(error)

  if (normalized.includes('abort') || normalized.includes('timeout')) {
    return new AppError('REQUEST_TIMEOUT', 'The request timed out', { endpoint, retryable: true })
  }
  if (status === 401 || status === 403 || /unauthorized|invalid.*(api|key)|api key/i.test(normalized)) {
    return new AppError('UNAUTHORIZED', 'The provider rejected the credentials', { status, endpoint })
  }
  if (status === 404 || /model.*(not found|does not exist)|unknown model/i.test(normalized)) {
    return new AppError('MODEL_NOT_FOUND', 'The selected model was not found', { status, endpoint })
  }
  if (status === 429 || /rate.?limit|too many requests|quota/i.test(normalized)) {
    return new AppError('RATE_LIMITED', 'The provider rate limit or quota was reached', {
      status,
      endpoint,
      retryable: true,
    })
  }
  if (/cors|cross.?origin|failed to fetch|networkerror|network request/i.test(normalized)) {
    return new AppError('CORS_BLOCKED', 'The browser blocked the provider request', { endpoint, retryable: true })
  }
  if (/budget|maximum.*(tool|call|duration)|recursion|loop|repeated identical tool/i.test(normalized)) {
    return new AppError(
      normalized.includes('budget') ? 'AGENT_BUDGET' : 'AGENT_LOOP',
      'The Agent exceeded its execution guard',
      {
        endpoint,
      },
    )
  }
  if (/conflict|changed since|changed before|hash/i.test(normalized)) {
    return new AppError('DOCUMENT_CONFLICT', 'The document changed before the edit was applied', { endpoint })
  }
  if (/content control|ooxml|table|image|drawing|non.?text/i.test(normalized)) {
    return new AppError('NON_TEXT_OBJECTS', 'The selected range contains non-text objects', { endpoint })
  }

  return new AppError('REQUEST_FAILED', message, { status, endpoint })
}

const SENSITIVE_KEY = /(api.?key|authorization|token|secret|password|cookie|credential)/i

export const sanitizeForLog = (value: unknown): unknown => {
  if (typeof value === 'string') {
    if (SENSITIVE_KEY.test(value)) return '[REDACTED]'
    return value.length > 240 ? `${value.slice(0, 240)}…` : value
  }
  if (Array.isArray(value)) return value.map(sanitizeForLog)
  if (value && typeof value === 'object') {
    return Object.fromEntries(
      Object.entries(value).map(([key, entry]) => [
        key,
        SENSITIVE_KEY.test(key) ? '[REDACTED]' : sanitizeForLog(entry),
      ]),
    )
  }
  return value
}

export const isDiagnosticsEnabled = () => localStorage.getItem('diagnosticsEnabled') === 'true'

export const safeLog = (event: string, details?: Record<string, unknown>) => {
  if (!isDiagnosticsEnabled()) return
  console.info(`[WordGPT] ${event}`, sanitizeForLog(details || {}))
}
