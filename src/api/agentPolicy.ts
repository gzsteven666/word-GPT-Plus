interface MessageLike {
  _getType?: () => string
}

export const resolveAgentToolChoice = (messages: MessageLike[]): 'required' | 'auto' =>
  messages.at(-1)?._getType?.() === 'human' ? 'required' : 'auto'
