import assert from 'node:assert/strict'

import { AppError, classifyError, sanitizeForLog } from '../src/api/errors.ts'

assert.equal(classifyError({ status: 401 }, 'https://api.example.com/v1/models').code, 'UNAUTHORIZED')
assert.equal(classifyError({ status: 429 }).code, 'RATE_LIMITED')
assert.equal(classifyError(new TypeError('Failed to fetch')).code, 'CORS_BLOCKED')
assert.equal(classifyError(new Error('Repeated identical tool call detected')).code, 'AGENT_LOOP')
assert.equal(classifyError(new Error('The selection changed before the edit was applied')).code, 'DOCUMENT_CONFLICT')
assert.equal(classifyError(new Error('maximum context length exceeded')).code, 'CONTEXT_TOO_LARGE')
assert.equal(classifyError({ status: 413 }).code, 'CONTEXT_TOO_LARGE')

const sanitized = sanitizeForLog({ apiKey: 'secret-value', authorization: 'Bearer secret-value', count: 2 }) as Record<
  string,
  unknown
>
assert.equal(sanitized.apiKey, '[REDACTED]')
assert.equal(sanitized.authorization, '[REDACTED]')
assert.equal(sanitized.count, 2)
assert.equal(new AppError('REQUEST_FAILED', 'failure').retryable, false)

console.log('errors tests: PASS')
