import assert from 'node:assert/strict'

import { resolveAgentToolChoice } from '../src/api/agentPolicy.ts'

const message = (type: string) => ({ _getType: () => type })

assert.equal(resolveAgentToolChoice([message('human')]), 'required')
assert.equal(resolveAgentToolChoice([message('human'), message('tool')]), 'auto')
assert.equal(resolveAgentToolChoice([message('human'), message('ai')]), 'auto')
assert.equal(resolveAgentToolChoice([]), 'auto')

console.log('agent policy tests: PASS')
