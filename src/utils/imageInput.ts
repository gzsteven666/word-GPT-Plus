import { HumanMessage, type Message } from '@langchain/core/messages'

export const SUPPORTED_IMAGE_TYPES = ['image/png', 'image/jpeg', 'image/webp'] as const
export const MAX_IMAGE_BYTES = 10 * 1024 * 1024
export const MAX_IMAGE_DIMENSION = 1600
export const MAX_IMAGE_PAYLOAD_BYTES = 4 * 1024 * 1024
export const MAX_IMAGE_ATTACHMENTS = 4
export const MAX_IMAGE_TOTAL_PAYLOAD_BYTES = 8 * 1024 * 1024

export type ImageCapabilityGate = 'allowed' | 'blocked' | 'unknown'

export type ImageInputErrorCode =
  | 'IMAGE_TYPE_UNSUPPORTED'
  | 'IMAGE_TOO_LARGE'
  | 'IMAGE_PAYLOAD_TOO_LARGE'
  | 'IMAGE_COUNT_EXCEEDED'
  | 'IMAGE_TOTAL_PAYLOAD_TOO_LARGE'

export class ImageInputError extends Error {
  readonly code: ImageInputErrorCode
  readonly fileName?: string

  constructor(code: ImageInputErrorCode, fileName?: string) {
    super(code)
    this.name = 'ImageInputError'
    this.code = code
    this.fileName = fileName
  }
}

export interface ImageAttachment {
  id: string
  name: string
  mimeType: (typeof SUPPORTED_IMAGE_TYPES)[number]
  size: number
  width: number
  height: number
  dataUrl: string
}

export const constrainImageDimensions = (width: number, height: number, maxDimension = MAX_IMAGE_DIMENSION) => {
  const scale = Math.min(1, maxDimension / Math.max(width, height))
  return {
    width: Math.max(1, Math.round(width * scale)),
    height: Math.max(1, Math.round(height * scale)),
  }
}

const textFromContent = (content: unknown): string => {
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

export const isSupportedImageFile = (file: File): boolean =>
  SUPPORTED_IMAGE_TYPES.includes(file.type as (typeof SUPPORTED_IMAGE_TYPES)[number]) && file.size <= MAX_IMAGE_BYTES

/** Extract image files from a clipboard without changing the browser's text-paste behavior. */
export const getClipboardImageFiles = (clipboardData: DataTransfer | null): File[] => {
  if (!clipboardData) return []

  const itemFiles = Array.from(clipboardData.items || [])
    .filter(item => item.kind === 'file')
    .map(item => item.getAsFile())
    .filter((file): file is File => Boolean(file))
  const files = itemFiles.length ? itemFiles : Array.from(clipboardData.files || [])
  return files.filter(file => SUPPORTED_IMAGE_TYPES.includes(file.type as (typeof SUPPORTED_IMAGE_TYPES)[number]))
}

export const getImageCapabilityGate = (capability: 'yes' | 'no' | 'unknown'): ImageCapabilityGate => {
  if (capability === 'yes') return 'allowed'
  if (capability === 'no') return 'blocked'
  return 'unknown'
}

/** Keep the durable/chat-visible copy text-only. The image request is built separately for one model call. */
export const sanitizeHistoryMessage = (message: Message): HumanMessage =>
  new HumanMessage(textFromContent(message.content))

export const buildEphemeralMultimodalMessage = (text: string, dataUrls: string | readonly string[]): HumanMessage => {
  const urls = typeof dataUrls === 'string' ? [dataUrls] : dataUrls
  return new HumanMessage({
    content: [
      { type: 'text', text },
      ...urls.map(dataUrl => ({ type: 'image_url' as const, image_url: { url: dataUrl } })),
    ],
  })
}

export const appendImageAttachments = (
  existing: readonly ImageAttachment[],
  incoming: readonly ImageAttachment[],
): ImageAttachment[] => {
  const combined = [...existing, ...incoming]
  if (combined.length > MAX_IMAGE_ATTACHMENTS) throw new ImageInputError('IMAGE_COUNT_EXCEEDED')
  if (combined.reduce((total, image) => total + image.size, 0) > MAX_IMAGE_TOTAL_PAYLOAD_BYTES) {
    throw new ImageInputError('IMAGE_TOTAL_PAYLOAD_TOO_LARGE')
  }
  return combined
}

export const clearSentImageAttachments = (
  current: readonly ImageAttachment[],
  sent: readonly ImageAttachment[],
  completed: boolean,
): ImageAttachment[] => {
  if (!completed) return [...current]
  const sentIds = new Set(sent.map(image => image.id))
  return current.filter(image => !sentIds.has(image.id))
}

const readAsDataUrl = (file: File): Promise<string> =>
  new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => resolve(String(reader.result || ''))
    reader.onerror = () => reject(reader.error || new Error('Unable to read image'))
    reader.readAsDataURL(file)
  })

const loadImage = (dataUrl: string): Promise<HTMLImageElement> =>
  new Promise((resolve, reject) => {
    const image = new Image()
    image.onload = () => resolve(image)
    image.onerror = () => reject(new Error('Unable to decode image'))
    image.src = dataUrl
  })

const canvasDataUrl = (image: HTMLImageElement, mimeType: string, quality: number): string => {
  const canvas = document.createElement('canvas')
  const dimensions = constrainImageDimensions(image.naturalWidth, image.naturalHeight)
  canvas.width = dimensions.width
  canvas.height = dimensions.height
  const context = canvas.getContext('2d')
  if (!context) throw new Error('Unable to process image')
  context.drawImage(image, 0, 0, canvas.width, canvas.height)
  return canvas.toDataURL(mimeType, quality)
}

const byteLengthOfDataUrl = (dataUrl: string): number => {
  const base64 = dataUrl.split(',', 2)[1] || ''
  return Math.ceil((base64.length * 3) / 4)
}

const prepareImageDataUrlAttachment = async (
  name: string,
  originalDataUrl: string,
  originalMimeType: string,
): Promise<ImageAttachment> => {
  if (!SUPPORTED_IMAGE_TYPES.includes(originalMimeType as (typeof SUPPORTED_IMAGE_TYPES)[number])) {
    throw new ImageInputError('IMAGE_TYPE_UNSUPPORTED', name)
  }
  const image = await loadImage(originalDataUrl)
  let dataUrl = originalDataUrl
  let mimeType = originalMimeType as (typeof SUPPORTED_IMAGE_TYPES)[number]

  if (
    byteLengthOfDataUrl(originalDataUrl) > MAX_IMAGE_PAYLOAD_BYTES ||
    image.naturalWidth > MAX_IMAGE_DIMENSION ||
    image.naturalHeight > MAX_IMAGE_DIMENSION
  ) {
    // JPEG keeps the request bounded for PNG screenshots while preserving the selected image in memory only.
    mimeType = originalMimeType === 'image/jpeg' ? 'image/jpeg' : 'image/webp'
    let quality = 0.86
    do {
      dataUrl = canvasDataUrl(image, mimeType, quality)
      quality -= 0.12
    } while (byteLengthOfDataUrl(dataUrl) > MAX_IMAGE_PAYLOAD_BYTES && quality >= 0.5)
    if (byteLengthOfDataUrl(dataUrl) > MAX_IMAGE_PAYLOAD_BYTES) {
      throw new ImageInputError('IMAGE_PAYLOAD_TOO_LARGE', name)
    }
  }

  return {
    id: globalThis.crypto?.randomUUID?.() || `${Date.now()}-${Math.random()}`,
    name,
    mimeType,
    size: byteLengthOfDataUrl(dataUrl),
    width: image.naturalWidth,
    height: image.naturalHeight,
    dataUrl,
  }
}

export async function prepareImageAttachment(file: File): Promise<ImageAttachment> {
  if (!SUPPORTED_IMAGE_TYPES.includes(file.type as (typeof SUPPORTED_IMAGE_TYPES)[number])) {
    throw new ImageInputError('IMAGE_TYPE_UNSUPPORTED', file.name)
  }
  if (file.size > MAX_IMAGE_BYTES) throw new ImageInputError('IMAGE_TOO_LARGE', file.name)
  return prepareImageDataUrlAttachment(file.name, await readAsDataUrl(file), file.type)
}

export const prepareImageAttachmentFromDataUrl = (
  name: string,
  dataUrl: string,
  mimeType: string,
): Promise<ImageAttachment> => prepareImageDataUrlAttachment(name, dataUrl, mimeType)
