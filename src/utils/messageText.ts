const TOOL_ACTIVITY_LINE = /^\s*(?:🔧\s*Calling tool:.*|✅\s*Tool .* completed)\s*$/gim

export const stripToolActivity = (text: string): string =>
  text
    .replace(TOOL_ACTIVITY_LINE, '')
    .replace(/\n{3,}/g, '\n\n')
    .trim()
