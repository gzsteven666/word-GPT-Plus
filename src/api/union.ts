import { BaseChatModel } from '@langchain/core/language_models/chat_models'
import { ChatGoogleGenerativeAI } from '@langchain/google-genai'
import { ChatGroq } from '@langchain/groq'
// import { MemorySaver } from '@langchain/langgraph'
import { ChatOllama } from '@langchain/ollama'
import { AzureChatOpenAI, ChatOpenAI } from '@langchain/openai'
import { createAgent, createMiddleware } from 'langchain'

import { resolveAgentToolChoice } from '@/api/agentPolicy'
import { IndexedDBSaver } from '@/api/checkpoints'
import { classifyError, safeLog } from '@/api/errors'

import {
  AgentOptions,
  AzureOptions,
  GeminiOptions,
  GroqOptions,
  OllamaOptions,
  OpenAIOptions,
  ProviderOptions,
} from './types'

const ModelCreators: Record<string, (opts: any) => BaseChatModel> = {
  official: (opts: OpenAIOptions) => {
    const modelName = opts.model || 'gpt-5'
    return new ChatOpenAI({
      modelName,
      configuration: {
        apiKey: opts.config.apiKey,
        baseURL: opts.config.baseURL || 'https://api.openai.com/v1',
        defaultHeaders: opts.config.headers,
      },
      temperature: opts.temperature,
      maxTokens: opts.maxTokens ?? 4096,
      timeout: opts.timeout,
    })
  },

  ollama: (opts: OllamaOptions) => {
    return new ChatOllama({
      model: opts.ollamaModel,
      baseUrl: opts.ollamaEndpoint?.replace(/\/$/, '') || 'http://localhost:11434',
      temperature: opts.temperature,
    })
  },

  groq: (opts: GroqOptions) => {
    return new ChatGroq({
      model: opts.groqModel,
      apiKey: opts.groqAPIKey,
      temperature: opts.temperature,
      maxTokens: opts.maxTokens ?? 4096,
      timeout: opts.timeout,
    })
  },

  gemini: (opts: GeminiOptions) => {
    return new ChatGoogleGenerativeAI({
      model: opts.geminiModel ?? 'gemini-3-pro-preview',
      apiKey: opts.geminiAPIKey,
      temperature: opts.temperature,
      maxOutputTokens: opts.maxTokens ?? 4096,
    })
  },

  azure: (opts: AzureOptions) => {
    return new AzureChatOpenAI({
      model: opts.azureDeploymentName,
      temperature: opts.temperature,
      maxTokens: opts.maxTokens ?? 4096,
      timeout: opts.timeout,
      azureOpenAIApiKey: opts.azureAPIKey,
      azureOpenAIEndpoint: opts.azureAPIEndpoint,
      azureOpenAIApiDeploymentName: opts.azureDeploymentName,
      azureOpenAIApiVersion: opts.azureAPIVersion ?? '2024-10-01',
    })
  },
}

// const checkpointer = new MemorySaver()
const checkpointer = new IndexedDBSaver()

const requireFirstToolCallMiddleware = createMiddleware({
  name: 'require-first-tool-call',
  wrapModelCall: (request, handler) =>
    handler({
      ...request,
      toolChoice: resolveAgentToolChoice(request.messages),
    }),
})

async function executeChatFlow(model: BaseChatModel, options: ProviderOptions): Promise<void> {
  try {
    const stream = await model.stream(options.messages, { signal: options.abortSignal })

    let fullContent = ''
    for await (const chunk of stream) {
      if (options.abortSignal?.aborted) break

      const content = Array.isArray(chunk.content)
        ? chunk.content
            .map(part =>
              typeof part === 'string' ? part : 'text' in part && typeof part.text === 'string' ? part.text : '',
            )
            .join('')
        : typeof chunk.content === 'string'
          ? chunk.content
          : ''
      fullContent += content
      options.onStream(fullContent)
    }
  } catch (error: any) {
    if (error.name === 'AbortError' || options.abortSignal?.aborted) {
      throw error
    }
    const classified = classifyError(error)
    options.errorIssue.value = classified.code
    safeLog('chat.error', { code: classified.code, status: classified.status })
  } finally {
    options.loading.value = false
  }
}

async function executeAgentFlow(model: BaseChatModel, options: AgentOptions): Promise<void> {
  try {
    if (!options.threadId) {
      options.threadId = crypto.randomUUID()
    }
    const agent = createAgent({
      model,
      tools: options.tools || [],
      checkpointer,
      middleware: options.forceToolCall ? [requireFirstToolCallMiddleware] : [],
    })

    const stream = await agent.stream(
      {
        messages: options.messages,
      },
      {
        recursionLimit: Number(options.recursionLimit), //最大迭代次数
        signal: options.abortSignal,
        configurable: {
          thread_id: options.threadId,
          checkpoint_id: options.checkpointId,
        },
        streamMode: 'values',
      },
    )

    let fullContent = ''
    let stepCount = 0
    let totalToolCalls = 0
    let totalModelCalls = 0
    let totalWrites = 0
    let totalExternalRequests = 0
    let estimatedCostUsd = 0
    let lastToolSignature = ''
    let repeatedToolCalls = 0
    const seenAiMessageIds = new Set<string>()
    const repeatedErrors = new Map<string, number>()
    const repeatedReads = new Map<string, number>()
    const startedAt = Date.now()

    for await (const step of stream) {
      if (options.abortSignal?.aborted) {
        break
      }

      stepCount++
      if (options.maxDurationMs && Date.now() - startedAt > options.maxDurationMs) {
        const budgetError = new Error('Agent maximum duration budget exceeded')
        budgetError.name = 'AgentBudgetError'
        throw budgetError
      }
      safeLog('agent.step', {
        step: stepCount,
        messageCount: step.messages?.length || 0,
        lastMessageType: step.messages?.[step.messages.length - 1]?.constructor?.name || 'unknown',
      })

      const messages = step.messages || []
      const lastMessage = messages[messages.length - 1]

      if (!lastMessage) continue

      // Cast to any for accessing tool-related properties
      const msg = lastMessage as any

      safeLog('agent.message', { type: msg._getType?.() || 'unknown' })

      if (msg._getType?.() === 'ai') {
        const messageId = String(msg.id || msg.lc_kwargs?.id || `step-${stepCount}`)
        if (!seenAiMessageIds.has(messageId)) {
          seenAiMessageIds.add(messageId)
          totalModelCalls++
          estimatedCostUsd += options.estimatedCostPerModelCallUsd || 0
          if (options.maxModelCalls && totalModelCalls > options.maxModelCalls) {
            const budgetError = new Error('Agent maximum model call budget exceeded')
            budgetError.name = 'AgentBudgetError'
            throw budgetError
          }
          if (options.maxCostUsd && estimatedCostUsd > options.maxCostUsd) {
            const budgetError = new Error('Agent estimated cost budget exceeded')
            budgetError.name = 'AgentBudgetError'
            throw budgetError
          }
        }
      }

      // Handle AI messages with tool calls
      if (msg._getType?.() === 'ai' && msg.tool_calls?.length > 0) {
        safeLog('agent.tool_calls', { count: msg.tool_calls.length })
        for (const toolCall of msg.tool_calls) {
          totalToolCalls++
          if (options.maxToolCalls && totalToolCalls > options.maxToolCalls) {
            const budgetError = new Error('Agent maximum tool call budget exceeded')
            budgetError.name = 'AgentBudgetError'
            throw budgetError
          }
          const isWrite = options.writeToolNames?.includes(toolCall.name) || false
          const isExternal = options.externalToolNames?.includes(toolCall.name) || false
          if (isWrite) {
            totalWrites++
            repeatedReads.clear()
            if (options.maxWrites && totalWrites > options.maxWrites) {
              const budgetError = new Error('Agent maximum write budget exceeded')
              budgetError.name = 'AgentBudgetError'
              throw budgetError
            }
          }
          if (isExternal) {
            totalExternalRequests++
            if (options.maxExternalRequests && totalExternalRequests > options.maxExternalRequests) {
              const budgetError = new Error('Agent maximum external request budget exceeded')
              budgetError.name = 'AgentBudgetError'
              throw budgetError
            }
          }
          const signature = `${toolCall.name}:${JSON.stringify(toolCall.args || {})}`
          if (!isWrite && /^(read_|getDocument|findText|search)/i.test(toolCall.name)) {
            const readCount = (repeatedReads.get(signature) || 0) + 1
            repeatedReads.set(signature, readCount)
            if (readCount >= 4) {
              const loopError = new Error('Agent repeatedly read the same region without making progress')
              loopError.name = 'AgentLoopError'
              throw loopError
            }
          }
          repeatedToolCalls = signature === lastToolSignature ? repeatedToolCalls + 1 : 1
          lastToolSignature = signature
          if (repeatedToolCalls >= 3) {
            const loopError = new Error('Repeated identical tool call detected')
            loopError.name = 'AgentLoopError'
            throw loopError
          }
          if (options.onToolCall) options.onToolCall(toolCall.name, toolCall.args)
        }
      }

      // Handle tool result messages
      if (msg._getType?.() === 'tool') {
        const toolName = msg.name || 'unknown'
        const toolContent = String(msg.content || '')
        safeLog('agent.tool_result', { name: toolName, contentLength: toolContent.length })
        if (/^Error:/i.test(toolContent)) {
          const category = classifyError(toolContent).code
          const errorCount = (repeatedErrors.get(category) || 0) + 1
          repeatedErrors.set(category, errorCount)
          if (errorCount >= 3) {
            const loopError = new Error(`Agent repeated the ${category} error category`)
            loopError.name = 'AgentLoopError'
            throw loopError
          }
        }
        if (options.onToolResult) {
          options.onToolResult(toolName, toolContent)
        }
      }

      // Handle AI message content (the final response)
      if (msg._getType?.() === 'ai' && msg.content) {
        const content = typeof msg.content === 'string' ? msg.content : ''
        if (content && (!msg.tool_calls || msg.tool_calls.length === 0)) {
          fullContent = content
          options.onStream(fullContent)
        }
      }
    }

    safeLog('agent.completed', {
      steps: stepCount,
      modelCalls: totalModelCalls,
      toolCalls: totalToolCalls,
      writes: totalWrites,
      externalRequests: totalExternalRequests,
      estimatedCostUsd,
    })
  } catch (error: any) {
    if (error.name === 'AbortError' || options.abortSignal?.aborted) {
      throw error
    }
    const classified = classifyError(error)
    options.errorIssue.value = error.name === 'AgentBudgetError' ? 'AGENT_BUDGET' : classified.code
    safeLog('agent.error', { code: options.errorIssue.value, status: classified.status })
  } finally {
    options.loading.value = false
  }
}

export async function getChatResponse(options: ProviderOptions) {
  const creator = ModelCreators[options.provider]
  if (!creator) {
    throw new Error(`Unsupported provider: ${options.provider}`)
  }
  const model = creator(options)
  return executeChatFlow(model, options)
}

export async function getAgentResponse(options: AgentOptions) {
  const creator = ModelCreators[options.provider]
  if (!creator) {
    throw new Error(`Unsupported provider: ${options.provider}`)
  }
  const model = creator(options)
  return executeAgentFlow(model, options)
}
