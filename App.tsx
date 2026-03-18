import {
  DndContext,
  KeyboardSensor,
  PointerSensor,
  closestCenter,
  useSensor,
  useSensors,
} from '@dnd-kit/core'
import type { DragEndEvent } from '@dnd-kit/core'
import {
  SortableContext,
  arrayMove,
  rectSortingStrategy,
  sortableKeyboardCoordinates,
  useSortable,
} from '@dnd-kit/sortable'
import { CSS } from '@dnd-kit/utilities'
import { PDFDocument, degrees } from 'pdf-lib'
import { GlobalWorkerOptions, getDocument } from 'pdfjs-dist'
import type { PDFDocumentProxy, PDFPageProxy } from 'pdfjs-dist'
import { useMemo, useRef, useState } from 'react'
import type { ChangeEvent, DragEvent } from 'react'
import pdfWorkerSrc from 'pdfjs-dist/build/pdf.worker.min.mjs?url'
import './App.css'

GlobalWorkerOptions.workerSrc = pdfWorkerSrc

const THUMBNAIL_WIDTH = 230
const PREVIEW_MAX_WIDTH = 1100

type SourceFile = {
  id: string
  name: string
  kind: 'pdf' | 'image'
  mime: string
  bytes: Uint8Array
}

type PageTile = {
  id: string
  sourceId: string
  sourceName: string
  pageIndex: number
  pageNumber: number
  thumbnail: string
  rotation: number
  kind: 'pdf' | 'image'
}

type ExportMode = 'original' | 'compressed'

function makeId() {
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
    return crypto.randomUUID()
  }

  return `${Date.now()}-${Math.random().toString(16).slice(2)}`
}

function pluralize(count: number, noun: string) {
  return `${count} ${noun}${count === 1 ? '' : 's'}`
}

function normalizeRotation(rotation: number) {
  const normalized = rotation % 360
  return normalized < 0 ? normalized + 360 : normalized
}

function clampNumber(value: number, min: number, max: number) {
  return Math.min(max, Math.max(min, value))
}

function parsePageRange(rangeText: string, totalPages: number) {
  if (totalPages < 1) {
    return []
  }

  const trimmed = rangeText.trim()

  if (!trimmed) {
    return Array.from({ length: totalPages }, (_, index) => index)
  }

  const selectedIndexes = new Set<number>()
  const segments = trimmed.split(',')

  for (const rawSegment of segments) {
    const segment = rawSegment.trim()

    if (!segment) {
      continue
    }

    if (/^\d+$/.test(segment)) {
      const pageNumber = Number(segment)

      if (pageNumber < 1 || pageNumber > totalPages) {
        throw new Error(
          `Page ${pageNumber} is outside the document range 1-${totalPages}.`,
        )
      }

      selectedIndexes.add(pageNumber - 1)
      continue
    }

    const rangeMatch = /^(\d*)-(\d*)$/.exec(segment)

    if (!rangeMatch) {
      throw new Error(`Invalid range "${segment}". Use values like 1-3,5,8-.`)
    }

    const startPage = rangeMatch[1] ? Number(rangeMatch[1]) : 1
    const endPage = rangeMatch[2] ? Number(rangeMatch[2]) : totalPages

    if (
      Number.isNaN(startPage) ||
      Number.isNaN(endPage) ||
      startPage < 1 ||
      endPage < 1 ||
      startPage > totalPages ||
      endPage > totalPages ||
      startPage > endPage
    ) {
      throw new Error(
        `Invalid range "${segment}" for document with ${totalPages} pages.`,
      )
    }

    for (let pageNumber = startPage; pageNumber <= endPage; pageNumber += 1) {
      selectedIndexes.add(pageNumber - 1)
    }
  }

  const parsedIndexes = [...selectedIndexes].sort((left, right) => left - right)

  if (parsedIndexes.length === 0) {
    throw new Error('Range did not match any pages.')
  }

  return parsedIndexes
}

function dataUrlToUint8Array(dataUrl: string) {
  const parts = dataUrl.split(',')

  if (parts.length < 2) {
    throw new Error('Invalid image data URL.')
  }

  const decoded = atob(parts[1])
  const byteArray = new Uint8Array(decoded.length)

  for (let index = 0; index < decoded.length; index += 1) {
    byteArray[index] = decoded.charCodeAt(index)
  }

  return byteArray
}

function rotateCanvas(sourceCanvas: HTMLCanvasElement, rotation: number) {
  const normalizedRotation = normalizeRotation(rotation)

  if (normalizedRotation === 0) {
    return sourceCanvas
  }

  const rotatedCanvas = document.createElement('canvas')
  const quarterTurn = normalizedRotation === 90 || normalizedRotation === 270

  rotatedCanvas.width = quarterTurn ? sourceCanvas.height : sourceCanvas.width
  rotatedCanvas.height = quarterTurn ? sourceCanvas.width : sourceCanvas.height

  const context = rotatedCanvas.getContext('2d')

  if (!context) {
    return sourceCanvas
  }

  context.translate(rotatedCanvas.width / 2, rotatedCanvas.height / 2)
  context.rotate((normalizedRotation * Math.PI) / 180)
  context.drawImage(
    sourceCanvas,
    -sourceCanvas.width / 2,
    -sourceCanvas.height / 2,
  )

  return rotatedCanvas
}

async function renderPageToCanvas(page: PDFPageProxy, scale: number) {
  const viewport = page.getViewport({ scale })
  const canvas = document.createElement('canvas')
  canvas.width = Math.max(1, Math.floor(viewport.width))
  canvas.height = Math.max(1, Math.floor(viewport.height))
  const context = canvas.getContext('2d')

  if (!context) {
    throw new Error('This browser cannot render PDF previews.')
  }

  await page.render({
    canvas,
    canvasContext: context,
    viewport,
  }).promise

  return canvas
}

async function openPdfDocumentWithPrompt(
  sourceBytes: Uint8Array,
  fileName: string,
  initialPassword?: string,
): Promise<{ doc: PDFDocumentProxy; passwordUsed?: string }> {
  let password = initialPassword

  for (let attempt = 0; attempt < 4; attempt += 1) {
    try {
      const loadingTask = getDocument({
        data: sourceBytes.slice(),
        ...(password ? { password } : {}),
      })
      const doc = await loadingTask.promise

      return { doc, passwordUsed: password }
    } catch (error) {
      const message = error instanceof Error ? error.message.toLowerCase() : ''
      const mayNeedPassword =
        message.includes('password') || message.includes('encrypted')

      if (!mayNeedPassword) {
        throw error
      }

      const promptMessage = password
        ? `${fileName}: incorrect password. Enter password again.`
        : `${fileName} is password protected. Enter its password.`
      const enteredPassword = window.prompt(promptMessage, '')

      if (!enteredPassword) {
        throw new Error(`${fileName}: password entry canceled.`)
      }

      password = enteredPassword
    }
  }

  throw new Error(`${fileName}: unable to unlock PDF after multiple attempts.`)
}

function wipeByteArray(bytes: Uint8Array) {
  try {
    bytes.fill(0)
  } catch {
    // Best-effort wipe only.
  }
}

function wipeSourceBytes(sources: SourceFile[]) {
  for (const source of sources) {
    wipeByteArray(source.bytes)
  }
}

function isPdfFile(file: File) {
  return (
    file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf')
  )
}

function isAcceptedImageFile(file: File) {
  return (
    /^image\//i.test(file.type) ||
    /\.(jpe?g|png|webp|gif|bmp|avif)$/i.test(file.name)
  )
}

async function renderImageToCanvas(
  bytes: Uint8Array,
  mime: string,
  targetWidth?: number,
): Promise<HTMLCanvasElement> {
  return new Promise((resolve, reject) => {
    const blob = new Blob([new Uint8Array(bytes)], { type: mime })
    const url = URL.createObjectURL(blob)
    const img = new Image()

    img.onload = () => {
      URL.revokeObjectURL(url)

      try {
        const scale = targetWidth != null ? targetWidth / img.naturalWidth : 1
        const canvas = document.createElement('canvas')
        canvas.width = Math.max(1, Math.floor(img.naturalWidth * scale))
        canvas.height = Math.max(1, Math.floor(img.naturalHeight * scale))
        const context = canvas.getContext('2d')

        if (!context) {
          reject(new Error('This browser cannot render image previews.'))
          return
        }

        context.drawImage(img, 0, 0, canvas.width, canvas.height)
        resolve(canvas)
      } catch (error) {
        reject(error instanceof Error ? error : new Error(String(error)))
      }
    }

    img.onerror = () => {
      URL.revokeObjectURL(url)
      reject(new Error('Failed to decode image.'))
    }

    img.src = url
  })
}

type SortablePageCardProps = {
  page: PageTile
  order: number
  onRemove: (pageId: string) => void
  onRotate: (pageId: string) => void
  onPreview: (page: PageTile) => void
}

function SortablePageCard({
  page,
  order,
  onRemove,
  onRotate,
  onPreview,
}: SortablePageCardProps) {
  const { attributes, listeners, setNodeRef, transform, transition, isDragging } =
    useSortable({
      id: page.id,
    })

  const style = {
    transform: CSS.Transform.toString(transform),
    transition,
  }

  return (
    <article
      ref={setNodeRef}
      className={`page-card${isDragging ? ' page-card--dragging' : ''}`}
      style={style}
    >
      <div className="page-card__preview">
        <button
          type="button"
          className="page-card__preview-button"
          onClick={() => onPreview(page)}
          aria-label={`Preview ${page.sourceName} page ${page.pageNumber}`}
        >
          <img
            src={page.thumbnail}
            alt={`${page.sourceName} page ${page.pageNumber}`}
            loading="lazy"
            draggable={false}
            style={{ transform: `rotate(${page.rotation}deg)` }}
          />
        </button>
      </div>

      <div className="page-card__meta">
        <p className="page-card__order">Position {order}</p>
        <p className="page-card__name" title={page.sourceName}>
          {page.sourceName}
        </p>
        <p className="page-card__page">
          {page.kind === 'image' ? 'Image' : `Page ${page.pageNumber}`}
        </p>
        {page.rotation > 0 ? (
          <p className="page-card__rotation">Rotated {page.rotation} deg</p>
        ) : (
          <p className="page-card__rotation">Tap Preview for full view</p>
        )}
      </div>

      <div className="page-card__controls">
        <button
          type="button"
          className="card-button card-button--preview"
          onClick={() => onPreview(page)}
          aria-label={`Preview page ${order}`}
        >
          Preview
        </button>
        <button
          type="button"
          className="card-button card-button--rotate"
          onClick={() => onRotate(page.id)}
          aria-label={`Rotate ${page.sourceName} page ${page.pageNumber}`}
        >
          Rotate +90
        </button>
        <button
          type="button"
          className="card-button card-button--drag"
          aria-label={`Drag page ${order}`}
          {...attributes}
          {...listeners}
        >
          Drag
        </button>
        <button
          type="button"
          className="card-button card-button--remove"
          onClick={() => onRemove(page.id)}
          aria-label={`Remove ${page.sourceName} page ${page.pageNumber}`}
        >
          Remove
        </button>
      </div>
    </article>
  )
}

function App() {
  const [sourcePdfs, setSourcePdfs] = useState<SourceFile[]>([])
  const [pageTiles, setPageTiles] = useState<PageTile[]>([])
  const [importRangeInput, setImportRangeInput] = useState('')
  const [exportMode, setExportMode] = useState<ExportMode>('original')
  const [rasterQuality, setRasterQuality] = useState(78)
  const [rasterScale, setRasterScale] = useState(1.4)
  const [autoClearAfterDownload, setAutoClearAfterDownload] = useState(true)
  const [isImporting, setIsImporting] = useState(false)
  const [isMerging, setIsMerging] = useState(false)
  const [dropActive, setDropActive] = useState(false)
  const [statusMessage, setStatusMessage] = useState(
    'Upload PDFs to start sorting individual pages.',
  )
  const [errorMessage, setErrorMessage] = useState<string | null>(null)
  const [previewTile, setPreviewTile] = useState<PageTile | null>(null)
  const [previewDataUrl, setPreviewDataUrl] = useState<string | null>(null)
  const [previewLoading, setPreviewLoading] = useState(false)
  const [previewError, setPreviewError] = useState<string | null>(null)
  const [isHelpOpen, setIsHelpOpen] = useState(false)
  const fileInputRef = useRef<HTMLInputElement>(null)
  const previewCacheRef = useRef<Map<string, string>>(new Map())
  const previewRequestRef = useRef<string>('')

  const sensors = useSensors(
    useSensor(PointerSensor, {
      activationConstraint: {
        distance: 6,
      },
    }),
    useSensor(KeyboardSensor, {
      coordinateGetter: sortableKeyboardCoordinates,
    }),
  )

  const sourceCount = useMemo(() => {
    return new Set(pageTiles.map((tile) => tile.sourceId)).size
  }, [pageTiles])

  const sourceById = useMemo(() => {
    return new Map(sourcePdfs.map((source) => [source.id, source]))
  }, [sourcePdfs])

  const usesRasterSettings = exportMode === 'compressed'

  const clearPreviewCacheForPage = (pageId: string) => {
    for (const cacheKey of Array.from(previewCacheRef.current.keys())) {
      if (cacheKey.startsWith(`${pageId}:`)) {
        previewCacheRef.current.delete(cacheKey)
      }
    }
  }

  const closePreview = () => {
    previewRequestRef.current = ''
    setPreviewTile(null)
    setPreviewDataUrl(null)
    setPreviewError(null)
    setPreviewLoading(false)
  }

  const purgeSensitiveData = (nextStatusMessage?: string) => {
    setSourcePdfs((currentSources) => {
      wipeSourceBytes(currentSources)
      return []
    })
    setPageTiles([])
    previewCacheRef.current.clear()
    closePreview()
    setErrorMessage(null)

    if (nextStatusMessage) {
      setStatusMessage(nextStatusMessage)
    }
  }

  const openPreview = async (tile: PageTile) => {
    const source = sourceById.get(tile.sourceId)

    if (!source) {
      setErrorMessage('Could not find the source file for preview.')
      return
    }

    const cacheKey = `${tile.id}:${tile.rotation}`
    const cachedPreview = previewCacheRef.current.get(cacheKey)

    setPreviewTile(tile)
    setPreviewError(null)

    if (cachedPreview) {
      setPreviewDataUrl(cachedPreview)
      setPreviewLoading(false)
      return
    }

    const requestId = makeId()
    previewRequestRef.current = requestId
    setPreviewDataUrl(null)
    setPreviewLoading(true)

    if (tile.kind === 'image') {
      try {
        const previewCanvas = await renderImageToCanvas(
          source.bytes,
          source.mime,
          PREVIEW_MAX_WIDTH,
        )
        const rotatedCanvas = rotateCanvas(previewCanvas, tile.rotation)
        const dataUrl = rotatedCanvas.toDataURL('image/jpeg', 0.92)

        if (previewRequestRef.current !== requestId) {
          return
        }

        previewCacheRef.current.set(cacheKey, dataUrl)
        setPreviewDataUrl(dataUrl)
      } catch (error) {
        const message =
          error instanceof Error ? error.message : 'Unable to render image preview.'

        if (previewRequestRef.current === requestId) {
          setPreviewError(message)
        }
      } finally {
        if (previewRequestRef.current === requestId) {
          setPreviewLoading(false)
        }
      }

      return
    }

    let previewDoc: PDFDocumentProxy | null = null

    try {
      const opened = await openPdfDocumentWithPrompt(source.bytes, source.name)
      previewDoc = opened.doc

      const page = await previewDoc.getPage(tile.pageNumber)

      try {
        const baseViewport = page.getViewport({ scale: 1 })
        const previewScale = clampNumber(PREVIEW_MAX_WIDTH / baseViewport.width, 1.1, 3)
        const renderedCanvas = await renderPageToCanvas(page, previewScale)
        const rotatedCanvas = rotateCanvas(renderedCanvas, tile.rotation)
        const dataUrl = rotatedCanvas.toDataURL('image/jpeg', 0.92)

        if (previewRequestRef.current !== requestId) {
          return
        }

        previewCacheRef.current.set(cacheKey, dataUrl)
        setPreviewDataUrl(dataUrl)
      } finally {
        page.cleanup()
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unable to render preview.'

      if (previewRequestRef.current === requestId) {
        setPreviewError(message)
      }
    } finally {
      if (previewDoc) {
        previewDoc.cleanup()
        await previewDoc.destroy()
      }

      if (previewRequestRef.current === requestId) {
        setPreviewLoading(false)
      }
    }
  }

  const clearAll = () => {
    purgeSensitiveData('All local file data was cleared from memory.')
  }

  const removeTile = (pageId: string) => {
    clearPreviewCacheForPage(pageId)

    if (previewTile?.id === pageId) {
      closePreview()
    }

    setPageTiles((currentTiles) => {
      const nextTiles = currentTiles.filter((tile) => tile.id !== pageId)
      const activeSourceIds = new Set(nextTiles.map((tile) => tile.sourceId))

      setSourcePdfs((currentSources) => {
        const retainedSources = currentSources.filter((source) => {
          return activeSourceIds.has(source.id)
        })
        const removedSources = currentSources.filter((source) => {
          return !activeSourceIds.has(source.id)
        })

        wipeSourceBytes(removedSources)

        return retainedSources
      })

      return nextTiles
    })
  }

  const rotateTile = (pageId: string) => {
    clearPreviewCacheForPage(pageId)

    if (previewTile?.id === pageId) {
      closePreview()
    }

    setPageTiles((currentTiles) => {
      return currentTiles.map((tile) => {
        if (tile.id !== pageId) {
          return tile
        }

        return {
          ...tile,
          rotation: normalizeRotation(tile.rotation + 90),
        }
      })
    })
  }

  const handleDragEnd = (event: DragEndEvent) => {
    const { active, over } = event

    if (!over || active.id === over.id) {
      return
    }

    setPageTiles((currentTiles) => {
      const oldIndex = currentTiles.findIndex((tile) => tile.id === active.id)
      const newIndex = currentTiles.findIndex((tile) => tile.id === over.id)

      if (oldIndex < 0 || newIndex < 0) {
        return currentTiles
      }

      return arrayMove(currentTiles, oldIndex, newIndex)
    })
  }

  const importFiles = async (incomingFiles: File[]) => {
    const validFiles = incomingFiles.filter(
      (file) => isPdfFile(file) || isAcceptedImageFile(file),
    )

    if (validFiles.length === 0) {
      setErrorMessage(
        'Please upload PDF or image files (JPEG, PNG, WebP, GIF, BMP, AVIF).',
      )
      return
    }

    setErrorMessage(null)
    setIsImporting(true)

    const nextSources: SourceFile[] = []
    const nextTiles: PageTile[] = []
    const parseErrors: string[] = []

    const selectedRangeText = importRangeInput.trim()

    for (const file of validFiles) {
      const sourceId = makeId()
      let sourceBytes: Uint8Array | null = null
      let keepSourceBytes = false

      if (isAcceptedImageFile(file) && !isPdfFile(file)) {
        try {
          setStatusMessage(`Loading image: ${file.name}...`)
          sourceBytes = new Uint8Array(await file.arrayBuffer())
          const mime = file.type || 'image/jpeg'
          const thumbnailCanvas = await renderImageToCanvas(
            sourceBytes,
            mime,
            THUMBNAIL_WIDTH,
          )

          nextSources.push({
            id: sourceId,
            name: file.name,
            kind: 'image',
            mime,
            bytes: sourceBytes,
          })
          keepSourceBytes = true

          nextTiles.push({
            id: makeId(),
            sourceId,
            sourceName: file.name,
            pageIndex: 0,
            pageNumber: 1,
            thumbnail: thumbnailCanvas.toDataURL('image/jpeg', 0.75),
            rotation: 0,
            kind: 'image',
          })
        } catch (error) {
          const message =
            error instanceof Error ? error.message : 'Unable to load this image.'
          parseErrors.push(`${file.name}: ${message}`)
        } finally {
          if (!keepSourceBytes && sourceBytes) {
            wipeByteArray(sourceBytes)
          }
        }
      } else {
        try {
          setStatusMessage(`Reading ${file.name}...`)

          sourceBytes = new Uint8Array(await file.arrayBuffer())
          const { doc: previewDoc } = await openPdfDocumentWithPrompt(
            sourceBytes,
            file.name,
          )

          try {
            const selectedIndexes = parsePageRange(
              selectedRangeText,
              previewDoc.numPages,
            )
            const fileTiles: PageTile[] = []

            for (
              let selectedIndex = 0;
              selectedIndex < selectedIndexes.length;
              selectedIndex += 1
            ) {
              const pageIndex = selectedIndexes[selectedIndex]
              const pageNumber = pageIndex + 1

              setStatusMessage(
                `Rendering ${file.name}: page ${selectedIndex + 1}/${selectedIndexes.length}`,
              )

              const page = await previewDoc.getPage(pageNumber)

              try {
                const baseViewport = page.getViewport({ scale: 1 })
                const scale = Math.min(1.25, THUMBNAIL_WIDTH / baseViewport.width)
                const renderedCanvas = await renderPageToCanvas(page, scale)

                fileTiles.push({
                  id: makeId(),
                  sourceId,
                  sourceName: file.name,
                  pageIndex,
                  pageNumber,
                  thumbnail: renderedCanvas.toDataURL('image/jpeg', 0.75),
                  rotation: 0,
                  kind: 'pdf',
                })
              } finally {
                page.cleanup()
              }
            }

            if (fileTiles.length === 0) {
              throw new Error('No readable pages were found in this file.')
            }

            nextSources.push({
              id: sourceId,
              name: file.name,
              kind: 'pdf',
              mime: 'application/pdf',
              bytes: sourceBytes,
            })
            keepSourceBytes = true

            nextTiles.push(...fileTiles)
          } finally {
            previewDoc.cleanup()
            await previewDoc.destroy()
          }
        } catch (error) {
          const message =
            error instanceof Error ? error.message : 'Unable to parse this PDF file.'
          parseErrors.push(`${file.name}: ${message}`)
        } finally {
          if (!keepSourceBytes && sourceBytes) {
            wipeByteArray(sourceBytes)
          }
        }
      }
    }

    setSourcePdfs((existingSources) => [...existingSources, ...nextSources])
    setPageTiles((existingTiles) => [...existingTiles, ...nextTiles])

    if (nextTiles.length > 0) {
      let summary =
        `Added ${pluralize(nextTiles.length, 'page')} from ${pluralize(nextSources.length, 'file')}. ` +
        'Drag cards to reorder and use Rotate +90 when needed.'

      if (selectedRangeText) {
        summary += ' Imported using your page range selection.'
      }

      if (nextSources.some((source) => source.kind === 'pdf')) {
        summary +=
          ' Password-protected PDFs may ask password again during preview or merge.'
      }

      setStatusMessage(summary)
    } else {
      setStatusMessage('No pages were added. Try another PDF or image file.')
    }

    if (parseErrors.length > 0) {
      setErrorMessage(
        `Skipped ${pluralize(parseErrors.length, 'file')}. First issue: ${parseErrors[0]}`,
      )
    }

    setIsImporting(false)
  }

  const onFileInputChange = (event: ChangeEvent<HTMLInputElement>) => {
    const selectedFiles = Array.from(event.target.files ?? [])
    void importFiles(selectedFiles)
    event.target.value = ''
  }

  const onDropZoneDragEnter = (event: DragEvent<HTMLElement>) => {
    event.preventDefault()
    setDropActive(true)
  }

  const onDropZoneDragOver = (event: DragEvent<HTMLElement>) => {
    event.preventDefault()
    event.dataTransfer.dropEffect = 'copy'
    setDropActive(true)
  }

  const onDropZoneDragLeave = (event: DragEvent<HTMLElement>) => {
    event.preventDefault()
    if (event.currentTarget.contains(event.relatedTarget as Node | null)) {
      return
    }

    setDropActive(false)
  }

  const onDropZoneDrop = (event: DragEvent<HTMLElement>) => {
    event.preventDefault()
    setDropActive(false)
    const droppedFiles = Array.from(event.dataTransfer.files)
    void importFiles(droppedFiles)
  }

  const rasterizeTileForMerge = async (
    tile: PageTile,
    source: SourceFile,
    rasterDocCache: Map<string, PDFDocumentProxy>,
    passwordCache: Map<string, string>,
  ) => {
    const mergeScale = clampNumber(rasterScale, 0.8, 2.4)
    const jpegQuality = clampNumber(rasterQuality / 100, 0.45, 0.95)

    let sourceDoc = rasterDocCache.get(source.id)

    if (!sourceDoc) {
      const passwordHint = passwordCache.get(source.id)
      const opened = await openPdfDocumentWithPrompt(
        source.bytes,
        source.name,
        passwordHint,
      )
      sourceDoc = opened.doc

      if (opened.passwordUsed) {
        passwordCache.set(source.id, opened.passwordUsed)
      }

      rasterDocCache.set(source.id, sourceDoc)
    }

    const page = await sourceDoc.getPage(tile.pageNumber)

    try {
      const renderedCanvas = await renderPageToCanvas(page, mergeScale)
      const rotatedCanvas = rotateCanvas(renderedCanvas, tile.rotation)
      const jpegBytes = dataUrlToUint8Array(
        rotatedCanvas.toDataURL('image/jpeg', jpegQuality),
      )

      return {
        jpegBytes,
        width: rotatedCanvas.width,
        height: rotatedCanvas.height,
      }
    } finally {
      page.cleanup()
    }
  }

  const mergeAndDownload = async () => {
    if (pageTiles.length === 0) {
      setErrorMessage('Add at least one page before merging.')
      return
    }

    setErrorMessage(null)
    setIsMerging(true)
    setStatusMessage('Building merged file...')

    const sourceLookup = new Map(sourcePdfs.map((source) => [source.id, source]))
    const vectorSourceCache = new Map<string, PDFDocument>()
    const rasterSourceCache = new Map<string, PDFDocumentProxy>()
    const mergePasswordCache = new Map<string, string>()
    let rasterizedPages = 0

    try {
      const mergedPdf = await PDFDocument.create()

      for (let index = 0; index < pageTiles.length; index += 1) {
        const tile = pageTiles[index]
        const source = sourceLookup.get(tile.sourceId)

        if (!source) {
          continue
        }

        if (tile.kind === 'image') {
          const imageCanvas = await renderImageToCanvas(source.bytes, source.mime)
          const rotatedCanvas = rotateCanvas(imageCanvas, tile.rotation)
          const jpegQuality = clampNumber(rasterQuality / 100, 0.45, 0.95)
          const jpegBytes = dataUrlToUint8Array(
            rotatedCanvas.toDataURL('image/jpeg', jpegQuality),
          )
          const embeddedImage = await mergedPdf.embedJpg(jpegBytes)
          const imagePage = mergedPdf.addPage([
            rotatedCanvas.width,
            rotatedCanvas.height,
          ])
          imagePage.drawImage(embeddedImage, {
            x: 0,
            y: 0,
            width: rotatedCanvas.width,
            height: rotatedCanvas.height,
          })
          setStatusMessage(`Merging ${index + 1}/${pageTiles.length} pages...`)
          continue
        }

        let usedVectorExport = false
        const shouldRasterize = exportMode === 'compressed'

        if (!shouldRasterize) {
          try {
            let loadedSource = vectorSourceCache.get(source.id)

            if (!loadedSource) {
              loadedSource = await PDFDocument.load(source.bytes)
              vectorSourceCache.set(source.id, loadedSource)
            }

            const [copiedPage] = await mergedPdf.copyPages(loadedSource, [tile.pageIndex])
            copiedPage.setRotation(degrees(tile.rotation))
            mergedPdf.addPage(copiedPage)
            usedVectorExport = true
          } catch {
            usedVectorExport = false
          }
        }

        if (!usedVectorExport) {
          const rasterizedPage = await rasterizeTileForMerge(
            tile,
            source,
            rasterSourceCache,
            mergePasswordCache,
          )
          const embeddedImage = await mergedPdf.embedJpg(rasterizedPage.jpegBytes)
          wipeByteArray(rasterizedPage.jpegBytes)
          const targetPage = mergedPdf.addPage([
            rasterizedPage.width,
            rasterizedPage.height,
          ])

          targetPage.drawImage(embeddedImage, {
            x: 0,
            y: 0,
            width: rasterizedPage.width,
            height: rasterizedPage.height,
          })

          rasterizedPages += 1
        }

        setStatusMessage(`Merging ${index + 1}/${pageTiles.length} pages...`)
      }

      if (mergedPdf.getPageCount() === 0) {
        throw new Error('No pages were available to export.')
      }

      const mergedBytes = await mergedPdf.save()
      const mergedArray = new Uint8Array(mergedBytes.length)
      mergedArray.set(mergedBytes)
      const mergedBlob = new Blob([mergedArray.buffer], {
        type: 'application/pdf',
      })
      wipeByteArray(mergedArray)
      const downloadUrl = URL.createObjectURL(mergedBlob)
      const timestamp = new Date().toISOString().slice(0, 19).replace(/[T:]/g, '-')
      const anchor = document.createElement('a')
      anchor.href = downloadUrl
      anchor.download = `merged-${timestamp}.pdf`
      document.body.append(anchor)
      anchor.click()
      anchor.remove()
      URL.revokeObjectURL(downloadUrl)

      const mergedCountMessage = `Merged ${pluralize(mergedPdf.getPageCount(), 'page')}. Download started.`

      if (rasterizedPages > 0) {
        if (autoClearAfterDownload) {
          purgeSensitiveData(
            `${mergedCountMessage} ${pluralize(rasterizedPages, 'page')} used image rendering. Local data auto-cleared.`,
          )
        } else {
          setStatusMessage(
            `${mergedCountMessage} ${pluralize(rasterizedPages, 'page')} used image rendering.`,
          )
        }
      } else if (autoClearAfterDownload) {
        purgeSensitiveData(`${mergedCountMessage} Local data auto-cleared.`)
      } else {
        setStatusMessage(mergedCountMessage)
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unable to create merged PDF.'
      setErrorMessage(message)
      setStatusMessage('Merge failed. Please try another set of files.')
    } finally {
      for (const rasterDoc of rasterSourceCache.values()) {
        rasterDoc.cleanup()
        await rasterDoc.destroy()
      }

      setIsMerging(false)
    }
  }

  return (
    <main className="app-shell">
      <section className="hero-panel">
        <button
          type="button"
          className="hero-panel__help"
          onClick={() => setIsHelpOpen(true)}
          aria-label="Open how this works guide"
          aria-haspopup="dialog"
          aria-expanded={isHelpOpen}
        >
          ?
        </button>

        <p className="hero-panel__eyebrow">PERGE</p>
        <h1>PDF Merger</h1>
        <p>
          Drop one or many PDFs, reorder every page visually, remove anything you do
          not need, then download a fresh combined file.
        </p>
      </section>

      <section
        className={`upload-panel${dropActive ? ' upload-panel--active' : ''}`}
        onDragEnter={onDropZoneDragEnter}
        onDragOver={onDropZoneDragOver}
        onDragLeave={onDropZoneDragLeave}
        onDrop={onDropZoneDrop}
      >
        <input
          ref={fileInputRef}
          type="file"
          accept="application/pdf,image/jpeg,image/png,image/webp,image/gif,image/bmp,image/avif"
          multiple
          className="upload-panel__input"
          onChange={onFileInputChange}
        />

        <button
          type="button"
          className="button button--primary"
          onClick={() => fileInputRef.current?.click()}
          disabled={isImporting || isMerging}
        >
          {isImporting ? 'Processing PDFs...' : 'Choose PDF files'}
        </button>

        <p className="upload-panel__hint">or drag and drop PDFs here</p>

        <label className="field upload-panel__config">
          <span>Import page range (applies to each uploaded file)</span>
          <input
            className="field__control"
            type="text"
            placeholder="Examples: 1-3,5,8-"
            value={importRangeInput}
            disabled={isImporting || isMerging}
            onChange={(event) => setImportRangeInput(event.target.value)}
          />
        </label>

        <div className="upload-panel__meta">
          <span>Mobile friendly drag-and-drop</span>
          <span>Page-level sorting</span>
          <span>Mix images with PDFs</span>
          <span>Password-protected PDFs</span>
        </div>
      </section>

      <section className="toolbar">
        <div className="toolbar__main">
          <div className="stats">
            <article className="stat-card">
              <strong>{sourceCount}</strong>
              <span>Files</span>
            </article>
            <article className="stat-card">
              <strong>{pageTiles.length}</strong>
              <span>Pages</span>
            </article>
          </div>

          <div className="export-settings">
            <label className="field">
              <span>Export mode</span>
              <select
                className="field__control"
                value={exportMode}
                onChange={(event) => setExportMode(event.target.value as ExportMode)}
                disabled={isMerging}
              >
                <option value="original">Original quality (vector)</option>
                <option value="compressed">Compressed (image based)</option>
              </select>
            </label>

            <label className="field field--slider">
              <span>Image quality: {rasterQuality}%</span>
              <input
                className="field__control"
                type="range"
                min={45}
                max={95}
                step={1}
                value={rasterQuality}
                disabled={!usesRasterSettings || isMerging}
                onChange={(event) => setRasterQuality(Number(event.target.value))}
              />
            </label>

            <label className="field field--slider">
              <span>Raster scale: {rasterScale.toFixed(1)}x</span>
              <input
                className="field__control"
                type="range"
                min={0.8}
                max={2.4}
                step={0.1}
                value={rasterScale}
                disabled={!usesRasterSettings || isMerging}
                onChange={(event) =>
                  setRasterScale(Number.parseFloat(event.target.value))
                }
              />
            </label>
          </div>

          <div className="privacy-controls">
            <label className="privacy-toggle">
              <input
                type="checkbox"
                checked={autoClearAfterDownload}
                disabled={isMerging}
                onChange={(event) => setAutoClearAfterDownload(event.target.checked)}
              />
              <span>Auto-clear files from memory after download</span>
            </label>

            <p className="settings-note" role="note">
              Privacy mode: files stay in-browser only, no cloud upload, no persistent
              storage. Password-protected files are asked again when needed.
            </p>
          </div>
        </div>

        <div className="toolbar__actions">
          <button
            type="button"
            className="button button--ghost"
            onClick={clearAll}
            disabled={pageTiles.length === 0 || isImporting || isMerging}
          >
            Clear all
          </button>
          <button
            type="button"
            className="button button--primary"
            onClick={() => void mergeAndDownload()}
            disabled={pageTiles.length === 0 || isImporting || isMerging}
          >
            {isMerging ? 'Merging...' : 'Merge and download'}
          </button>
        </div>
      </section>

      <p className="status-message" role="status">
        {statusMessage}
      </p>

      {errorMessage ? (
        <p className="error-message" role="alert">
          {errorMessage}
        </p>
      ) : null}

      {pageTiles.length === 0 ? (
        <section className="empty-state" aria-live="polite">
          <h2>No pages loaded yet</h2>
          <p>
            Add PDFs or images and this area will become a sortable board. Drag any
            card to reorder before merging.
          </p>
        </section>
      ) : (
        <DndContext
          sensors={sensors}
          collisionDetection={closestCenter}
          onDragEnd={handleDragEnd}
        >
          <SortableContext
            items={pageTiles.map((tile) => tile.id)}
            strategy={rectSortingStrategy}
          >
            <section className="page-grid" aria-label="Sortable PDF pages">
              {pageTiles.map((tile, index) => {
                return (
                  <SortablePageCard
                    key={tile.id}
                    page={tile}
                    order={index + 1}
                    onRemove={removeTile}
                    onRotate={rotateTile}
                    onPreview={(page) => {
                      void openPreview(page)
                    }}
                  />
                )
              })}
            </section>
          </SortableContext>
        </DndContext>
      )}

      {isHelpOpen ? (
        <section
          className="help-modal"
          role="dialog"
          aria-modal="true"
          aria-labelledby="help-modal-title"
        >
          <button
            type="button"
            className="help-modal__backdrop"
            onClick={() => setIsHelpOpen(false)}
            aria-label="Close guide"
          />

          <article className="help-modal__panel">
            <header className="help-modal__header">
              <div>
                <h2 id="help-modal-title">How this PDF merge process works</h2>
                <p>Quick guide for new users</p>
              </div>

              <button
                type="button"
                className="help-modal__close"
                onClick={() => setIsHelpOpen(false)}
              >
                Close
              </button>
            </header>

            <div className="help-modal__body">
              <ol className="help-modal__list">
                <li>
                  Upload PDFs and/or images (JPEG, PNG, WebP, GIF, BMP, AVIF). You
                  can mix them freely. For PDFs, optionally set a page range first,
                  for example 1-3,5,8-.
                </li>
                <li>
                  Each selected page appears as a card. Drag cards to reorder, rotate
                  pages by 90 deg, remove pages, or open large preview to inspect
                  content.
                </li>
                <li>
                  Choose export mode. Original quality preserves vector pages when
                  possible. Compressed mode converts pages to images to reduce size.
                </li>
                <li>
                  Click Merge and download to create the final PDF in your chosen
                  page order.
                </li>
              </ol>

              <p className="help-modal__privacy">
                Privacy note: files are processed in-browser, not uploaded to a server.
                You can keep Auto-clear enabled to remove in-memory data after
                download.
              </p>
            </div>
          </article>
        </section>
      ) : null}

      {previewTile ? (
        <section className="preview-modal" role="dialog" aria-modal="true">
          <button
            type="button"
            className="preview-modal__backdrop"
            onClick={closePreview}
            aria-label="Close preview"
          />

          <article className="preview-modal__panel">
            <header className="preview-modal__header">
              <div>
                <h2>{previewTile.sourceName}</h2>
                <p>
                  Page {previewTile.pageNumber}
                  {previewTile.rotation > 0
                    ? ` - Rotated ${previewTile.rotation} deg`
                    : ''}
                </p>
              </div>

              <button
                type="button"
                className="preview-modal__close"
                onClick={closePreview}
              >
                Close
              </button>
            </header>

            <div className="preview-modal__body">
              {previewLoading ? (
                <p className="preview-modal__loading">Rendering large preview...</p>
              ) : null}

              {previewError ? (
                <p className="preview-modal__error" role="alert">
                  {previewError}
                </p>
              ) : null}

              {previewDataUrl ? (
                <img
                  src={previewDataUrl}
                  alt={`${previewTile.sourceName} page ${previewTile.pageNumber} full preview`}
                  className="preview-modal__image"
                />
              ) : null}
            </div>
          </article>
        </section>
      ) : null}
    </main>
  )
}

export default App
