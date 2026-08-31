export type WordImageReadErrorCode =
  | 'WORD_IMAGE_NOT_SELECTED'
  | 'WORD_IMAGE_AMBIGUOUS'
  | 'WORD_IMAGE_UNSUPPORTED'
  | 'WORD_IMAGE_READ_FAILED'

export class WordImageReadError extends Error {
  readonly code: WordImageReadErrorCode

  constructor(code: WordImageReadErrorCode, message: string) {
    super(message)
    this.name = 'WordImageReadError'
    this.code = code
  }
}

export interface WordSelectedInlineImage {
  dataUrl: string
  width: number
  height: number
  altTextTitle: string
  altTextDescription: string
  imageFormat: string
}

const imageFormatToMime = (format: string): string => {
  const normalized = format.toLowerCase()
  if (normalized === 'png') return 'image/png'
  if (normalized === 'jpeg' || normalized === 'jpg') return 'image/jpeg'
  if (normalized === 'webp') return 'image/webp'
  throw new WordImageReadError('WORD_IMAGE_UNSUPPORTED', `Word returned an unsupported image format: ${format}`)
}

const toDataUrl = (value: string, format: string): string => {
  if (value.startsWith('data:image/')) return value
  return `data:${imageFormatToMime(format)};base64,${value}`
}

export const readSelectedInlineImage = async (): Promise<WordSelectedInlineImage> => {
  try {
    return await Word.run(async context => {
      const selection = context.document.getSelection()
      const pictures = selection.inlinePictures
      pictures.load('items')
      await context.sync()

      if (pictures.items.length === 0) {
        throw new WordImageReadError('WORD_IMAGE_NOT_SELECTED', 'No inline image is selected')
      }
      if (pictures.items.length > 1) {
        throw new WordImageReadError('WORD_IMAGE_AMBIGUOUS', 'Select exactly one inline image')
      }

      const picture = pictures.items[0]
      picture.load('width,height,altTextTitle,altTextDescription,imageFormat')
      const imageResult = picture.getBase64ImageSrc()
      await context.sync()

      const value = String(imageResult.value || '')
      if (!value) throw new WordImageReadError('WORD_IMAGE_READ_FAILED', 'Word returned an empty image')

      return {
        dataUrl: toDataUrl(value, String(picture.imageFormat || '')),
        width: Number(picture.width) || 0,
        height: Number(picture.height) || 0,
        altTextTitle: String(picture.altTextTitle || ''),
        altTextDescription: String(picture.altTextDescription || ''),
        imageFormat: String(picture.imageFormat || ''),
      }
    })
  } catch (error) {
    if (error instanceof WordImageReadError) throw error
    throw new WordImageReadError('WORD_IMAGE_READ_FAILED', error instanceof Error ? error.message : String(error))
  }
}
