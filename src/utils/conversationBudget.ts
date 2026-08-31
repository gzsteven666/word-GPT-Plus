import type { BaseMessage } from '@langchain/core/messages'

/** Conservative client-side ceiling to avoid sending an overlong chat history. */
export const DEFAULT_MAX_CONTEXT_CHARS = 48_000

const contentToText = (content: BaseMessage['content']): string => {
  if (typeof content === 'string') return content
  if (!Array.isArray(content)) return ''

  return content
    .map(part => {
      if (typeof part === 'string') return part
      if (part && typeof part === 'object' && 'text' in part && typeof part.text === 'string') return part.text
      return ''
    })
    .join('')
}

export const estimateMessageChars = (messages: readonly Pick<BaseMessage, 'content'>[]): number =>
  messages.reduce((total, message) => total + contentToText(message.content).length, 0)

export const exceedsContextBudget = (
  messages: readonly Pick<BaseMessage, 'content'>[],
  maxChars = DEFAULT_MAX_CONTEXT_CHARS,
): boolean => estimateMessageChars(messages) > maxChars
