import {
  Children,
  forwardRef,
  isValidElement,
  memo,
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
} from 'react'
import {
  AlertTriangle,
  Check,
  Copy,
  Download,
  Eye,
  EyeOff,
  MessageSquare,
  Moon,
  PanelLeft,
  Paperclip,
  Send,
  Settings,
  Sun,
  Trash2,
  X,
} from 'lucide-react'
import { renderToStaticMarkup } from 'react-dom/server'
import katexCss from 'katex/dist/katex.min.css?inline'
import { VariableSizeList as List } from 'react-window'
import { marked } from 'marked'
import ReactMarkdown from 'react-markdown'
import remarkGfm from 'remark-gfm'
import remarkMath from 'remark-math'
import rehypeKatex from 'rehype-katex'
import { Prism as SyntaxHighlighter } from 'react-syntax-highlighter'
import * as XLSX from 'xlsx'
import 'katex/dist/katex.min.css'

const PROVIDERS = {
  anthropic: {
    id: 'anthropic',
    name: 'Anthropic',
    shortName: 'Claude',
    color: '#C48A5A',
    models: [
      { id: 'claude-sonnet-4-20250514', name: 'Claude Sonnet 4' },
      { id: 'claude-opus-4-20250514', name: 'Claude Opus 4' },
      { id: 'claude-haiku-3-5-20241022', name: 'Claude 3.5 Haiku' },
    ],
    tempRange: { min: 0, max: 1, default: 1, step: 0.1 },
    maxTokensRange: { min: 1, max: 8192, default: 1024 },
    keyPrefix: 'sk-ant-',
    endpoint: 'https://api.anthropic.com/v1/messages',
  },
  openai: {
    id: 'openai',
    name: 'OpenAI',
    shortName: 'GPT',
    color: '#10A37F',
    models: [
      { id: 'gpt-5.2', name: 'GPT-5.2' },
      { id: 'gpt-5.2-pro-2025-12-11', name: 'GPT-5.2 Pro (2025-12-11)' },
      { id: 'gpt-5.2-pro', name: 'GPT-5.2 Pro' },
      { id: 'gpt-5.2-codex', name: 'GPT-5.2 Codex' },
      { id: 'gpt-5.2-chat-latest', name: 'GPT-5.2 Chat Latest' },
      { id: 'gpt-5.2-2025-12-11', name: 'GPT-5.2 (2025-12-11)' },
      { id: 'gpt-4o', name: 'GPT-4o' },
      { id: 'gpt-4o-mini', name: 'GPT-4o Mini' },
      { id: 'gpt-4-turbo', name: 'GPT-4 Turbo' },
      { id: 'gpt-4.1', name: 'GPT-4.1' },
      { id: 'gpt-4.1-mini', name: 'GPT-4.1 Mini' },
      { id: 'gpt-4.1-nano', name: 'GPT-4.1 Nano' },
      { id: 'gpt-4.1-nano-2025-04-14', name: 'GPT-4.1 Nano (2025-04-14)' },
      { id: 'gpt-4.1-mini-2025-04-14', name: 'GPT-4.1 Mini (2025-04-14)' },
      { id: 'gpt-4.1-2025-04-14', name: 'GPT-4.1 (2025-04-14)' },
      { id: 'o3', name: 'o3' },
      { id: 'o4-mini', name: 'o4-mini' },
      { id: 'o1-pro', name: 'o1-pro' },
      { id: 'o1', name: 'o1' },
      { id: 'o1-2024-12-17', name: 'o1 (2024-12-17)' },
      { id: 'o1-pro-2025-03-19', name: 'o1-pro (2025-03-19)' },
      { id: 'o3-2025-04-16', name: 'o3 (2025-04-16)' },
      { id: 'o3-deep-research', name: 'o3 Deep Research' },
      { id: 'o3-mini', name: 'o3 Mini' },
      { id: 'o3-mini-2025-01-31', name: 'o3 Mini (2025-01-31)' },
      { id: 'o3-pro', name: 'o3 Pro' },
      { id: 'o3-pro-2025-06-10', name: 'o3 Pro (2025-06-10)' },
      { id: 'o4-mini-2025-04-16', name: 'o4-mini (2025-04-16)' },
      { id: 'o4-mini-deep-research', name: 'o4-mini Deep Research' },
      { id: 'o4-mini-deep-research-2025-06-26', name: 'o4-mini Deep Research (2025-06-26)' },
      { id: 'gpt-3.5-turbo', name: 'GPT-3.5 Turbo' },
    ],
    tempRange: { min: 0, max: 2, default: 1, step: 0.1 },
    maxTokensRange: { min: 1, max: 50000, default: 50000 },
    keyPrefix: 'sk-',
    endpoint: 'https://api.openai.com/v1/chat/completions',
  },
  gemini: {
    id: 'gemini',
    name: 'Google Gemini',
    shortName: 'Gemini',
    color: '#1A73E8',
    models: [
      { id: 'gemini-2.0-flash', name: 'Gemini 2.0 Flash' },
      { id: 'gemini-1.5-pro', name: 'Gemini 1.5 Pro' },
      { id: 'gemini-1.5-flash', name: 'Gemini 1.5 Flash' },
    ],
    tempRange: { min: 0, max: 2, default: 1, step: 0.1 },
    maxTokensRange: { min: 1, max: 8192, default: 1024 },
    keyPrefix: 'AI',
    endpoint: 'https://generativelanguage.googleapis.com/v1beta/models',
  },
  grok: {
    id: 'grok',
    name: 'xAI Grok',
    shortName: 'Grok',
    color: '#111111',
    models: [
      { id: 'grok-2', name: 'Grok 2' },
      { id: 'grok-beta', name: 'Grok Beta' },
    ],
    tempRange: { min: 0, max: 2, default: 1, step: 0.1 },
    maxTokensRange: { min: 1, max: 4096, default: 1024 },
    keyPrefix: 'xai-',
    endpoint: 'https://api.x.ai/v1/chat/completions',
  },
  mistral: {
    id: 'mistral',
    name: 'Mistral AI',
    shortName: 'Mistral',
    color: '#FF7000',
    models: [
      { id: 'mistral-large-latest', name: 'Mistral Large' },
      { id: 'mistral-small-latest', name: 'Mistral Small' },
      { id: 'codestral-latest', name: 'Codestral' },
    ],
    tempRange: { min: 0, max: 1, default: 0.7, step: 0.1 },
    maxTokensRange: { min: 1, max: 8192, default: 1024 },
    keyPrefix: '',
    endpoint: 'https://api.mistral.ai/v1/chat/completions',
  },
  cohere: {
    id: 'cohere',
    name: 'Cohere',
    shortName: 'Cohere',
    color: '#6B5CE7',
    models: [
      { id: 'command-r-plus', name: 'Command R+' },
      { id: 'command-r', name: 'Command R' },
      { id: 'command', name: 'Command' },
    ],
    tempRange: { min: 0, max: 1, default: 0.7, step: 0.1 },
    maxTokensRange: { min: 1, max: 4096, default: 1024 },
    keyPrefix: '',
    endpoint: 'https://api.cohere.ai/v1/chat',
  },
}

const PROVIDER_ORDER = [
  'anthropic',
  'openai',
  'gemini',
  'grok',
  'mistral',
  'cohere',
]

const DEFAULT_SYSTEM_PROMPT = 'You are a helpful AI assistant.'

const canUseStorage = () =>
  typeof window !== 'undefined' && typeof window.localStorage !== 'undefined'

const loadStored = (key, fallback) => {
  if (!canUseStorage()) {
    return fallback
  }
  try {
    const raw = window.localStorage.getItem(key)
    if (raw === null) {
      return fallback
    }
    return JSON.parse(raw)
  } catch (error) {
    return fallback
  }
}

const saveStored = (key, value) => {
  if (!canUseStorage()) {
    return
  }
  try {
    window.localStorage.setItem(key, JSON.stringify(value))
  } catch (error) {
    // Ignore write errors (storage disabled, quota, etc.)
  }
}

const clamp = (value, min, max) => Math.min(max, Math.max(min, value))

const STREAMING_PROVIDERS = new Set(['openai', 'grok', 'mistral'])
const VIRTUAL_ESTIMATE = 160
const VIRTUAL_GAP = 16

const normalizeMathDelimiters = (value) => {
  if (!value) {
    return ''
  }
  let result = ''
  let index = 0
  let inFence = false
  let inInline = false
  while (index < value.length) {
    if (!inInline && value.startsWith('```', index)) {
      inFence = !inFence
      result += '```'
      index += 3
      continue
    }
    const currentChar = value[index]
    if (!inFence && currentChar === '`') {
      inInline = !inInline
      result += currentChar
      index += 1
      continue
    }
    if (!inFence && !inInline && value.startsWith('\\[', index)) {
      result += '$$'
      index += 2
      continue
    }
    if (!inFence && !inInline && value.startsWith('\\]', index)) {
      result += '$$'
      index += 2
      continue
    }
    if (!inFence && !inInline && value.startsWith('\\(', index)) {
      result += '$'
      index += 2
      continue
    }
    if (!inFence && !inInline && value.startsWith('\\)', index)) {
      result += '$'
      index += 2
      continue
    }
    result += currentChar
    index += 1
  }
  return result
}

const normalizeMathForClipboard = (value) => {
  const normalized = normalizeMathDelimiters(value)
  let result = ''
  let index = 0
  let inFence = false
  let inInline = false
  while (index < normalized.length) {
    if (!inInline && normalized.startsWith('```', index)) {
      inFence = !inFence
      result += '```'
      index += 3
      continue
    }
    const currentChar = normalized[index]
    if (!inFence && currentChar === '`') {
      inInline = !inInline
      result += currentChar
      index += 1
      continue
    }
    if (!inFence && !inInline && normalized.startsWith('$$', index)) {
      const endIndex = normalized.indexOf('$$', index + 2)
      if (endIndex !== -1) {
        const content = normalized
          .slice(index + 2, endIndex)
          .replace(/\s+/g, ' ')
          .trim()
        result += `$${content}$`
        index = endIndex + 2
        continue
      }
    }
    result += currentChar
    index += 1
  }
  return result
}

const ListOuter = forwardRef(({ style, ...props }, ref) => (
  <div
    ref={ref}
    style={{ ...style }}
    className="h-full min-h-0 overflow-y-auto px-6 py-5"
    {...props}
  />
))

ListOuter.displayName = 'ListOuter'

const VirtualRow = memo(({ index, style, data }) => {
  const rowRef = useRef(null)
  const item = data.items[index]

  useLayoutEffect(() => {
    const element = rowRef.current
    if (!element) {
      return
    }
    const measure = () => {
      const height = element.getBoundingClientRect().height
      if (height) {
        data.setSize(index, height)
      }
    }
    measure()
    if (typeof ResizeObserver === 'undefined') {
      return
    }
    const observer = new ResizeObserver(() => measure())
    observer.observe(element)
    return () => observer.disconnect()
  }, [data, index, item])

  return (
    <div style={{ ...style, width: '100%' }}>
      <div ref={rowRef} style={{ paddingBottom: VIRTUAL_GAP }}>
        {item?.type === 'typing' ? (
          <TypingIndicator providerId={data.providerId} />
        ) : (
          <MessageBubble message={item} onError={data.onError} />
        )}
      </div>
    </div>
  )
})

VirtualRow.displayName = 'VirtualRow'

const extractStreamDelta = (eventName, payload) => {
  if (!payload) {
    return ''
  }
  if (
    payload.type === 'response.output_text.delta' &&
    typeof payload.delta === 'string'
  ) {
    return payload.delta
  }
  if (
    eventName === 'response.output_text.delta' &&
    typeof payload.delta === 'string'
  ) {
    return payload.delta
  }
  if (payload.choices?.[0]?.delta?.content) {
    return payload.choices[0].delta.content
  }
  if (payload.choices?.[0]?.text) {
    return payload.choices[0].text
  }
  if (typeof payload.delta === 'string') {
    return payload.delta
  }
  return ''
}

const generateId = () => {
  if (typeof crypto !== 'undefined' && crypto.randomUUID) {
    return crypto.randomUUID()
  }
  return `id-${Date.now()}-${Math.random().toString(16).slice(2)}`
}

const copyToClipboard = async (text) => {
  if (!text) {
    return false
  }
  if (navigator?.clipboard && window.isSecureContext) {
    try {
      await navigator.clipboard.writeText(text)
      return true
    } catch (error) {
      // Fallback to legacy copy.
    }
  }
  try {
    const textarea = document.createElement('textarea')
    textarea.value = text
    textarea.setAttribute('readonly', '')
    textarea.style.position = 'fixed'
    textarea.style.left = '-9999px'
    textarea.style.top = '0'
    textarea.style.opacity = '0'
    document.body.appendChild(textarea)
    textarea.focus()
    textarea.select()
    textarea.setSelectionRange(0, textarea.value.length)
    const success = document.execCommand('copy')
    textarea.remove()
    return success
  } catch (error) {
    return false
  }
}

const formatTime = (timestamp) => {
  try {
    return new Date(timestamp).toLocaleTimeString([], {
      hour: '2-digit',
      minute: '2-digit',
    })
  } catch (error) {
    return ''
  }
}

const formatFileSize = (size) => {
  if (!Number.isFinite(size)) {
    return ''
  }
  if (size < 1024) {
    return `${size} B`
  }
  if (size < 1024 * 1024) {
    return `${(size / 1024).toFixed(1)} KB`
  }
  if (size < 1024 * 1024 * 1024) {
    return `${(size / (1024 * 1024)).toFixed(1)} MB`
  }
  return `${(size / (1024 * 1024 * 1024)).toFixed(1)} GB`
}

const getModelName = (providerId, modelId) => {
  const provider = PROVIDERS[providerId]
  const match = provider?.models?.find((model) => model.id === modelId)
  return match ? match.name : modelId
}

const getKeyWarning = (providerId, apiKey) => {
  const prefix = PROVIDERS[providerId]?.keyPrefix
  if (!apiKey || !prefix) {
    return ''
  }
  return apiKey.startsWith(prefix)
    ? ''
    : `Expected a key starting with ${prefix}`
}

const deriveChatTitle = (messages) => {
  const firstUser = messages.find(
    (message) => message.role === 'user' && message.content?.trim(),
  )
  if (!firstUser) {
    return 'New chat'
  }
  const trimmed = (firstUser.displayContent || firstUser.content).trim()
  return trimmed.length > 36 ? `${trimmed.slice(0, 36)}...` : trimmed
}

const createChat = (messages = []) => {
  const now = new Date().toISOString()
  return {
    id: generateId(),
    title: deriveChatTitle(messages),
    messages,
    createdAt: now,
    updatedAt: now,
  }
}

const normalizeChat = (chat) => {
  const messages = Array.isArray(chat?.messages) ? chat.messages : []
  const createdAt = chat?.createdAt || new Date().toISOString()
  const updatedAt = chat?.updatedAt || createdAt
  return {
    id: chat?.id || generateId(),
    title: chat?.title || deriveChatTitle(messages),
    messages,
    createdAt,
    updatedAt,
  }
}

const getChatPreview = (messages) => {
  if (!messages.length) {
    return 'No messages yet'
  }
  const lastMessage = messages[messages.length - 1]
  const previewContent = lastMessage?.displayContent || lastMessage?.content
  if (!previewContent) {
    return 'No messages yet'
  }
  const trimmed = previewContent.trim()
  return trimmed.length > 60 ? `${trimmed.slice(0, 60)}...` : trimmed
}

const REASONING_PREVIEW_START = '<reasoning_preview>'
const REASONING_PREVIEW_END = '</reasoning_preview>'
const REASONING_PREVIEW_INSTRUCTION =
  'Provide a brief, high-level reasoning preview (1-2 sentences) wrapped between ' +
  `${REASONING_PREVIEW_START} and ${REASONING_PREVIEW_END}. ` +
  'Do not reveal chain-of-thought or step-by-step reasoning. If not possible, omit the tags.'

const extractReasoningPreview = (content) => {
  if (!content) {
    return { preview: '', cleaned: content }
  }
  const startIndex = content.indexOf(REASONING_PREVIEW_START)
  const endIndex = content.indexOf(REASONING_PREVIEW_END)
  if (startIndex === -1 || endIndex === -1 || endIndex <= startIndex) {
    return { preview: '', cleaned: content }
  }
  const preview = content
    .slice(startIndex + REASONING_PREVIEW_START.length, endIndex)
    .trim()
  const cleaned = `${content.slice(0, startIndex)}${content.slice(
    endIndex + REASONING_PREVIEW_END.length,
  )}`.trim()
  return { preview, cleaned }
}

const MAX_ATTACHMENT_CHARS = 120000

const resolveFileKind = (file) => {
  const name = file.name.toLowerCase()
  if (name.endsWith('.pdf')) {
    return 'pdf'
  }
  if (name.endsWith('.csv')) {
    return 'csv'
  }
  if (name.endsWith('.xlsx') || name.endsWith('.xls')) {
    return 'excel'
  }
  if (name.endsWith('.txt')) {
    return 'text'
  }
  if (file.type.includes('pdf')) {
    return 'pdf'
  }
  if (file.type.includes('csv')) {
    return 'csv'
  }
  if (file.type.includes('spreadsheet') || file.type.includes('excel')) {
    return 'excel'
  }
  if (file.type.startsWith('text/')) {
    return 'text'
  }
  return 'text'
}

const truncateText = (text) => {
  if (text.length <= MAX_ATTACHMENT_CHARS) {
    return text
  }
  return `${text.slice(0, MAX_ATTACHMENT_CHARS)}\n...[truncated]`
}

const readPdfText = async (file) => {
  const pdfjsLib = await import('pdfjs-dist/legacy/build/pdf')
  pdfjsLib.GlobalWorkerOptions.workerSrc = new URL(
    'pdfjs-dist/legacy/build/pdf.worker.min.mjs',
    import.meta.url,
  ).toString()
  const data = await file.arrayBuffer()
  const pdf = await pdfjsLib.getDocument({ data }).promise
  const pages = []
  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber)
    const content = await page.getTextContent()
    const pageText = content.items
      .map((item) => item.str)
      .join(' ')
      .replace(/\s+/g, ' ')
      .trim()
    if (pageText) {
      pages.push(`Page ${pageNumber}:\n${pageText}`)
    }
  }
  return truncateText(pages.join('\n\n'))
}

const readExcelText = async (file) => {
  const data = await file.arrayBuffer()
  const workbook = XLSX.read(data, { type: 'array' })
  const sheets = workbook.SheetNames.map((name) => {
    const sheet = workbook.Sheets[name]
    const csv = XLSX.utils.sheet_to_csv(sheet)
    return `Sheet: ${name}\n${csv}`
  })
  return truncateText(sheets.join('\n\n'))
}

const readAttachmentText = async (file) => {
  const kind = resolveFileKind(file)
  if (kind === 'pdf') {
    return { text: await readPdfText(file), kind }
  }
  if (kind === 'excel') {
    return { text: await readExcelText(file), kind }
  }
  const text = await file.text()
  return { text: truncateText(text), kind }
}

const formatAttachmentForPrompt = (attachment) => {
  const safeText = attachment.text?.trim()
  const label = `[Attachment: ${attachment.name} (${attachment.kind})]`
  if (!safeText) {
    return `${label}\n(No text could be extracted.)`
  }
  const language =
    attachment.kind === 'csv'
      ? 'csv'
      : attachment.kind === 'excel'
        ? 'csv'
        : 'text'
  return `${label}\n\`\`\`${language}\n${safeText}\n\`\`\``
}

const normalizeVerbosity = (value) =>
  ['low', 'medium', 'high'].includes(value) ? value : 'high'

const normalizeReasoningEffort = (model, effort) => {
  const allowed = ['none', 'low', 'medium', 'high']
  const selected = allowed.includes(effort) ? effort : 'none'
  if (!model) {
    return selected
  }
  if (model.startsWith('gpt-5.2-pro')) {
    if (selected === 'none' || selected === 'low') {
      return 'medium'
    }
    return selected
  }
  if (model.startsWith('gpt-5.2-codex')) {
    if (selected === 'none') {
      return 'low'
    }
    return selected
  }
  return selected
}

const buildResponsesInput = (messages, systemPrompt) => {
  const input = []
  const trimmedSystem = systemPrompt?.trim()
  if (trimmedSystem) {
    input.push({
      role: 'system',
      content: [{ type: 'input_text', text: trimmedSystem }],
    })
  }
  messages.forEach((message) => {
    const isAssistant = message.role === 'assistant'
    input.push({
      role: isAssistant ? 'assistant' : 'user',
      content: [
        {
          type: isAssistant ? 'output_text' : 'input_text',
          text: message.content,
        },
      ],
    })
  })
  return input
}

const extractTextValue = (value) => {
  if (typeof value === 'string') {
    return value.trim()
  }
  if (value && typeof value === 'object') {
    if (typeof value.value === 'string') {
      return value.value.trim()
    }
    if (typeof value.text === 'string') {
      return value.text.trim()
    }
    if (typeof value.content === 'string') {
      return value.content.trim()
    }
  }
  return ''
}

const parseResponsesText = (data) => {
  const directOutput =
    extractTextValue(data?.output_text) ||
    extractTextValue(data?.response?.output_text)
  if (directOutput) {
    return directOutput
  }

  const output = Array.isArray(data?.output)
    ? data.output
    : Array.isArray(data?.response?.output)
      ? data.response.output
      : []

  for (const item of output) {
    if (item?.type === 'message') {
      const content = Array.isArray(item?.content) ? item.content : []
      for (const contentItem of content) {
        const contentText =
          extractTextValue(contentItem?.output_text) ||
          extractTextValue(contentItem?.text) ||
          extractTextValue(contentItem?.content)
        if (contentText) {
          return contentText
        }
        const refusalText = extractTextValue(contentItem?.refusal)
        if (refusalText) {
          return refusalText
        }
      }
      continue
    }

    const itemText =
      extractTextValue(item?.output_text) ||
      extractTextValue(item?.text) ||
      extractTextValue(item?.content)
    if (itemText) {
      return itemText
    }

    const content = Array.isArray(item?.content) ? item.content : []
    for (const contentItem of content) {
      const contentText =
        extractTextValue(contentItem?.output_text) ||
        extractTextValue(contentItem?.text) ||
        extractTextValue(contentItem?.content)
      if (contentText) {
        return contentText
      }
      const refusalText = extractTextValue(contentItem?.refusal)
      if (refusalText) {
        return refusalText
      }
    }
  }

  return ''
}

const sanitizeFilename = (value) =>
  value.replace(/[<>:"/\\|?*\u0000-\u001F]/g, '').slice(0, 80) || 'response'

const formatExportTimestamp = () =>
  new Date().toISOString().replace('T', ' ').replace(/[:]/g, '-').slice(0, 19)

const getResponseTitle = (markdown) => {
  const normalized = normalizeMathDelimiters(markdown || '')
  const tokens = marked.lexer(normalized)
  const heading = tokens.find((token) => token.type === 'heading' && token.text)
  if (heading?.text) {
    return heading.text.trim()
  }
  const paragraph = tokens.find(
    (token) => token.type === 'paragraph' && token.text,
  )
  if (paragraph?.text) {
    return paragraph.text.trim()
  }
  const fallbackLine = normalized
    .split('\n')
    .find((line) => line.trim())
  return fallbackLine ? fallbackLine.trim() : 'response'
}

const buildExportMarkdownComponents = () => ({
  p: ({ children }) => <p style={{ margin: '0 0 10px' }}>{children}</p>,
  strong: ({ children }) => <strong>{children}</strong>,
  em: ({ children }) => <em>{children}</em>,
  ul: ({ children }) => <ul style={{ margin: '0 0 10px 18px' }}>{children}</ul>,
  ol: ({ children }) => <ol style={{ margin: '0 0 10px 18px' }}>{children}</ol>,
  li: ({ children }) => <li style={{ marginBottom: '6px' }}>{children}</li>,
  a: ({ href, children }) => (
    <a href={href} style={{ color: '#1d4ed8', textDecoration: 'underline' }}>
      {children}
    </a>
  ),
  code: ({ inline, className, children }) => {
    const text = String(children ?? '').replace(/\n$/, '')
    if (inline) {
      return (
        <code
          style={{
            background: '#f1f5f9',
            padding: '1px 4px',
            borderRadius: '4px',
            fontFamily: 'Consolas, Menlo, monospace',
            fontSize: '12px',
          }}
        >
          {text}
        </code>
      )
    }
    return (
      <pre
        style={{
          background: '#0f172a',
          color: '#f8fafc',
          padding: '12px',
          borderRadius: '10px',
          overflow: 'auto',
        }}
      >
        <code className={className}>{text}</code>
      </pre>
    )
  },
  pre: ({ children }) => <>{children}</>,
  blockquote: ({ children }) => (
    <blockquote
      style={{
        borderLeft: '3px solid #e2e8f0',
        paddingLeft: '12px',
        color: '#475569',
        margin: '12px 0',
      }}
    >
      {children}
    </blockquote>
  ),
  h1: ({ children }) => <h1 style={{ margin: '18px 0 8px' }}>{children}</h1>,
  h2: ({ children }) => <h2 style={{ margin: '16px 0 8px' }}>{children}</h2>,
  h3: ({ children }) => <h3 style={{ margin: '14px 0 8px' }}>{children}</h3>,
  hr: () => <hr style={{ margin: '14px 0', borderColor: '#e2e8f0' }} />,
  table: ({ children }) => (
    <table
      style={{
        width: '100%',
        borderCollapse: 'collapse',
        margin: '12px 0',
        fontSize: '12px',
      }}
    >
      {children}
    </table>
  ),
  th: ({ children }) => (
    <th
      style={{
        border: '1px solid #e2e8f0',
        padding: '6px 8px',
        textAlign: 'left',
        background: '#f1f5f9',
      }}
    >
      {children}
    </th>
  ),
  td: ({ children }) => (
    <td style={{ border: '1px solid #e2e8f0', padding: '6px 8px' }}>
      {children}
    </td>
  ),
})

const renderMarkdownToHtml = (markdown) => {
  const normalized = normalizeMathDelimiters(markdown || '')
  return renderToStaticMarkup(
    <ReactMarkdown
      remarkPlugins={[remarkGfm, remarkMath]}
      rehypePlugins={[rehypeKatex]}
      components={buildExportMarkdownComponents()}
    >
      {normalized}
    </ReactMarkdown>,
  )
}

const buildExportHtml = (markdown) => `
  <style>
    ${katexCss}
    .markdown-export-root {
      font-family: "Georgia", "Palatino Linotype", "Times New Roman", serif;
      color: #0f172a;
      line-height: 1.7;
      font-size: 15px;
      max-width: 720px;
      margin: 0 auto;
      padding: 6px 6px 36px;
    }
    .markdown-export-root .katex {
      font-size: 1em;
    }
    .markdown-export-root .katex-display {
      margin: 12px 0;
    }
    .markdown-export-root ul,
    .markdown-export-root ol {
      margin: 0 0 12px 20px;
      padding-left: 16px;
      list-style-position: outside;
    }
    .markdown-export-root ul {
      list-style-type: disc;
    }
    .markdown-export-root ol {
      list-style-type: decimal;
    }
    .markdown-export-root li {
      margin-bottom: 6px;
    }
    .markdown-export-root li::marker {
      color: #0f172a;
    }
    .markdown-export-root p,
    .markdown-export-root li,
    .markdown-export-root blockquote,
    .markdown-export-root pre,
    .markdown-export-root table,
    .markdown-export-root h1,
    .markdown-export-root h2,
    .markdown-export-root h3,
    .markdown-export-root h4,
    .markdown-export-root h5,
    .markdown-export-root h6 {
      break-inside: avoid;
      page-break-inside: avoid;
    }
  </style>
  <div class="markdown-export-root">${renderMarkdownToHtml(markdown)}</div>
`

const exportMarkdownAsPdf = async (markdown, filename) => {
  const html2pdfModule = await import('html2pdf.js')
  const html2pdf = html2pdfModule.default || html2pdfModule

  const container = document.createElement('div')
  container.style.position = 'fixed'
  container.style.left = '0'
  container.style.top = '0'
  container.style.right = '0'
  container.style.bottom = '0'
  container.style.zIndex = '9999'
  container.style.overflow = 'auto'
  container.style.background = '#ffffff'
  container.style.padding = '24px'
  container.style.color = '#0f172a'
  container.innerHTML = buildExportHtml(markdown)
  document.body.appendChild(container)

  if (document.fonts?.ready) {
    await document.fonts.ready
  }
  await new Promise((resolve) => {
    requestAnimationFrame(() => requestAnimationFrame(resolve))
  })
  await new Promise((resolve) => {
    setTimeout(resolve, 120)
  })

  const renderedText = container.textContent?.trim()
  if (!renderedText) {
    throw new Error('Rendered document is empty.')
  }

  const contentRoot =
    container.querySelector('.markdown-export-root') || container
  const exportWidth = Math.ceil(contentRoot.scrollWidth || 800)
  const exportHeight = Math.ceil(contentRoot.scrollHeight || 600)

  await html2pdf()
    .set({
      margin: 36,
      filename: `${sanitizeFilename(filename)}.pdf`,
      pagebreak: {
        mode: ['css', 'legacy'],
        avoid: [
          'p',
          'li',
          'blockquote',
          'pre',
          'table',
          'h1',
          'h2',
          'h3',
          'h4',
          'h5',
          'h6',
        ],
      },
      html2canvas: {
        scale: 2,
        backgroundColor: '#ffffff',
        useCORS: true,
        scrollX: 0,
        scrollY: 0,
        windowWidth: exportWidth,
        windowHeight: exportHeight,
      },
      jsPDF: { unit: 'pt', format: 'a4', orientation: 'portrait' },
    })
    .from(contentRoot)
    .save()

  container.remove()
}

const downloadBlob = (blob, filename) => {
  const url = URL.createObjectURL(blob)
  const link = document.createElement('a')
  link.href = url
  link.download = filename
  document.body.appendChild(link)
  link.click()
  link.remove()
  URL.revokeObjectURL(url)
}

const extractTextContent = (node) => {
  if (node === null || node === undefined) {
    return ''
  }
  if (typeof node === 'string' || typeof node === 'number') {
    return String(node)
  }
  if (Array.isArray(node)) {
    return node.map(extractTextContent).join('')
  }
  if (isValidElement(node)) {
    return extractTextContent(node.props?.children)
  }
  return ''
}

const collectRows = (node) => {
  const rows = []
  const children = Children.toArray(node?.props?.children ?? [])
  children.forEach((child) => {
    if (!isValidElement(child)) {
      return
    }
    if (child.type === 'tr') {
      const cells = Children.toArray(child.props?.children ?? [])
        .filter((cell) => isValidElement(cell))
        .map((cell) => extractTextContent(cell.props?.children).trim())
      rows.push(cells)
      return
    }
    rows.push(...collectRows(child))
  })
  return rows
}

const extractTableData = (tableChildren) => {
  const sections = Children.toArray(tableChildren ?? [])
  let headers = []
  let rows = []
  sections.forEach((section) => {
    if (!isValidElement(section)) {
      return
    }
    if (section.type === 'thead') {
      const headRows = collectRows(section)
      if (headRows.length) {
        headers = headRows[0]
      }
      return
    }
    if (section.type === 'tbody') {
      rows = rows.concat(collectRows(section))
      return
    }
    if (section.type === 'tr') {
      rows = rows.concat(collectRows(section))
      return
    }
    if (section.props?.children) {
      rows = rows.concat(collectRows(section))
    }
  })
  if (!headers.length && rows.length) {
    headers = rows[0]
    rows = rows.slice(1)
  }
  return { headers, rows }
}

const toCsv = (headers, rows) => {
  const escapeCell = (value) => {
    const text = String(value ?? '')
    const escaped = text.replace(/"/g, '""')
    return `"${escaped}"`
  }
  const lines = []
  if (headers.length) {
    lines.push(headers.map(escapeCell).join(','))
  }
  rows.forEach((row) => {
    lines.push(row.map(escapeCell).join(','))
  })
  return lines.join('\n')
}

const monokaiSublimeTheme = {
  'code[class*="language-"]': {
    color: '#F8F8F2',
    fontFamily:
      '"SFMono-Regular", Consolas, "Liberation Mono", Menlo, monospace',
    fontSize: '0.75rem',
    lineHeight: '1.6',
  },
  'pre[class*="language-"]': {
    color: '#F8F8F2',
    background: '#272822',
    borderRadius: '0.75rem',
    padding: '12px',
    overflow: 'auto',
  },
  comment: { color: '#75715E' },
  punctuation: { color: '#F8F8F2' },
  property: { color: '#A6E22E' },
  tag: { color: '#F92672' },
  boolean: { color: '#AE81FF' },
  number: { color: '#AE81FF' },
  constant: { color: '#AE81FF' },
  symbol: { color: '#66D9EF' },
  selector: { color: '#F92672' },
  'attr-name': { color: '#A6E22E' },
  string: { color: '#E6DB74' },
  char: { color: '#E6DB74' },
  builtin: { color: '#66D9EF' },
  operator: { color: '#F8F8F2' },
  entity: { color: '#F8F8F2' },
  url: { color: '#E6DB74' },
  variable: { color: '#F8F8F2' },
  'atrule': { color: '#F92672' },
  'attr-value': { color: '#E6DB74' },
  function: { color: '#A6E22E' },
  'class-name': { color: '#A6E22E' },
  keyword: { color: '#F92672' },
  regex: { color: '#FD971F' },
  important: { color: '#FD971F', fontWeight: 'bold' },
  deleted: { color: '#F92672' },
  inserted: { color: '#A6E22E' },
}

const isSoloMarkdownCodeBlock = (content) => {
  const trimmed = content?.trim()
  if (!trimmed || !trimmed.startsWith('```') || !trimmed.endsWith('```')) {
    return false
  }
  const fenceMatches = trimmed.match(/```/g)
  return fenceMatches?.length === 2
}

const MarkdownCodeBlock = ({ inline, className, children, isSolo }) => {
  const [copied, setCopied] = useState(false)
  const codeText = String(children ?? '').replace(/\n$/, '')
  const language =
    typeof className === 'string'
      ? className.replace('language-', '').trim()
      : ''

  const handleCopy = async () => {
    if (!codeText) {
      return
    }
    const success = await copyToClipboard(codeText)
    if (success) {
      setCopied(true)
      setTimeout(() => setCopied(false), 1200)
    } else {
      // Ignore clipboard errors.
    }
  }

  if (inline) {
    return (
    <code className="rounded bg-slate-100 px-1 py-0.5 text-[0.85em] text-slate-700 dark:bg-[#414141] dark:text-[#ffffff]">
      {children}
    </code>
  )
  }

  return (
    <div className="relative mt-2">
      <div className="absolute right-2 top-2 flex items-center gap-2">
        {language ? (
          <span className="rounded-full bg-black/40 px-2 py-0.5 text-[10px] uppercase tracking-wide text-white/70">
            {language}
          </span>
        ) : null}
        <button
          type="button"
          onClick={handleCopy}
          aria-label="Copy code"
          className="rounded-full bg-white/10 p-1 text-white/80 backdrop-blur transition hover:bg-white/20"
        >
          <Copy className={`h-3.5 w-3.5 ${copied ? 'text-emerald-300' : ''}`} />
        </button>
      </div>
      <SyntaxHighlighter
        language={language || 'text'}
        style={monokaiSublimeTheme}
        customStyle={{
          margin: 0,
          background: isSolo ? '#272822' : '#0F172A',
          borderRadius: '0.75rem',
          padding: '12px',
        }}
      >
        {codeText}
      </SyntaxHighlighter>
    </div>
  )
}

const createMarkdownComponents = (isSoloCodeBlock) => ({
  p: ({ children }) => (
    <p className="mb-2 last:mb-0 whitespace-pre-wrap">{children}</p>
  ),
  strong: ({ children }) => <strong className="font-semibold">{children}</strong>,
  em: ({ children }) => <em className="italic">{children}</em>,
  ul: ({ children }) => <ul className="ml-4 list-disc space-y-1">{children}</ul>,
  ol: ({ children }) => (
    <ol className="ml-4 list-decimal space-y-1">{children}</ol>
  ),
  li: ({ children }) => <li className="leading-relaxed">{children}</li>,
  a: ({ href, children }) => (
    <a
      href={href}
      target="_blank"
      rel="noreferrer"
      className="text-blue-600 underline decoration-transparent underline-offset-2 transition hover:decoration-blue-600 dark:text-blue-400 dark:hover:decoration-blue-300"
    >
      {children}
    </a>
  ),
  code: ({ inline, className, children }) => (
    <MarkdownCodeBlock
      inline={inline}
      className={className}
      isSolo={isSoloCodeBlock}
    >
      {children}
    </MarkdownCodeBlock>
  ),
  pre: ({ children }) => <>{children}</>,
  blockquote: ({ children }) => (
    <blockquote className="border-l-2 border-slate-300 pl-3 text-slate-600 dark:border-[#fff3] dark:text-[#cdcdcd]">
      {children}
    </blockquote>
  ),
  h1: ({ children }) => <h1 className="text-base font-semibold">{children}</h1>,
  h2: ({ children }) => <h2 className="text-sm font-semibold">{children}</h2>,
  h3: ({ children }) => <h3 className="text-sm font-semibold">{children}</h3>,
  hr: () => <hr className="my-3 border-slate-200 dark:border-[#ffffff26]" />,
  table: ({ children }) => {
    const { headers, rows } = extractTableData(children)
    const hasContent = headers.length > 0 || rows.length > 0

    const handleExportCsv = () => {
      if (!hasContent) {
        return
      }
      const csv = toCsv(headers, rows)
      downloadBlob(
        new Blob([csv], { type: 'text/csv;charset=utf-8;' }),
        'table.csv',
      )
    }

    const handleExportXlsx = () => {
      if (!hasContent) {
        return
      }
      const sheetData = headers.length ? [headers, ...rows] : rows
      const worksheet = XLSX.utils.aoa_to_sheet(sheetData)
      const workbook = XLSX.utils.book_new()
      XLSX.utils.book_append_sheet(workbook, worksheet, 'Table')
      XLSX.writeFile(workbook, 'table.xlsx', { compression: true })
    }

    return (
      <div className="my-3 overflow-x-auto rounded-2xl border border-slate-200 bg-white/70 p-3 shadow-sm dark:border-[#ffffff26] dark:bg-[#303030]/70">
        <div className="mb-2 flex items-center justify-between text-xs text-slate-500 dark:text-[#afafaf]">
          <span>Table</span>
          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={handleExportCsv}
              className="rounded-full border border-slate-200 px-2 py-0.5 text-[11px] font-semibold text-slate-600 transition hover:border-slate-300 hover:text-slate-900 dark:border-[#ffffff26] dark:text-[#cdcdcd] dark:hover:border-[#fff3] dark:hover:text-[#ffffff]"
            >
              CSV
            </button>
            <button
              type="button"
              onClick={handleExportXlsx}
              className="rounded-full border border-slate-200 px-2 py-0.5 text-[11px] font-semibold text-slate-600 transition hover:border-slate-300 hover:text-slate-900 dark:border-[#ffffff26] dark:text-[#cdcdcd] dark:hover:border-[#fff3] dark:hover:text-[#ffffff]"
            >
              XLSX
            </button>
          </div>
        </div>
        <table className="min-w-full border-collapse text-left text-xs">
          {children}
        </table>
      </div>
    )
  },
  thead: ({ children }) => (
    <thead className="border-b border-slate-200 dark:border-[#ffffff26]">
      {children}
    </thead>
  ),
  th: ({ children }) => (
    <th className="border border-slate-200 px-2 py-1 font-semibold text-slate-700 dark:border-[#ffffff26] dark:text-[#f3f3f3]">
      {children}
    </th>
  ),
  td: ({ children }) => (
    <td className="border border-slate-200 px-2 py-1 text-slate-700 dark:border-[#ffffff26] dark:text-[#f3f3f3]">
      {children}
    </td>
  ),
})

const openAiCompatibleAdapter = {
  formatRequest: (messages, settings) => {
    const system = settings.systemPrompt?.trim()
    const mapped = messages.map((message) => ({
      role: message.role,
      content: message.content,
    }))
    const isGpt5 = settings.model?.startsWith('gpt-5')
    if (isGpt5) {
      const systemWithPreview =
        settings.reasoningPreviewEnabled && system
          ? `${system}\n\n${REASONING_PREVIEW_INSTRUCTION}`
          : settings.reasoningPreviewEnabled
            ? REASONING_PREVIEW_INSTRUCTION
            : system
      const input = buildResponsesInput(messages, systemWithPreview)
      const normalizedEffort = normalizeReasoningEffort(
        settings.model,
        settings.reasoningEffort,
      )
      const normalizedVerbosity = normalizeVerbosity(settings.verbosity)
      const allowTemperature =
        !settings.model?.startsWith('gpt-5.2') || normalizedEffort === 'none'
      return {
        model: settings.model,
        input,
        max_output_tokens: settings.maxTokens,
        ...(allowTemperature ? { temperature: settings.temperature } : {}),
        reasoning: { effort: normalizedEffort },
        text: { verbosity: normalizedVerbosity },
      }
    }
    const isO3 = settings.model?.startsWith('o3')
    return {
      model: settings.model,
      ...(isO3
        ? { max_completion_tokens: settings.maxTokens }
        : { max_tokens: settings.maxTokens }),
      temperature: settings.temperature,
      messages: system
        ? [{ role: 'system', content: system }, ...mapped]
        : mapped,
    }
  },
  getHeaders: (apiKey) => ({
    'Content-Type': 'application/json',
    Authorization: `Bearer ${apiKey}`,
  }),
  getEndpoint: (model, providerEndpoint) =>
    model?.startsWith('gpt-5')
      ? 'https://api.openai.com/v1/responses'
      : providerEndpoint,
  parseResponse: (data) =>
    data?.choices?.[0]?.message?.content ??
    data?.choices?.[0]?.text ??
    parseResponsesText(data),
}

const ADAPTERS = {
  anthropic: {
    formatRequest: (messages, settings) => ({
      model: settings.model,
      max_tokens: settings.maxTokens,
      temperature: settings.temperature,
      system: settings.systemPrompt,
      messages: messages.map((message) => ({
        role: message.role,
        content: message.content,
      })),
    }),
    getHeaders: (apiKey) => ({
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
      'anthropic-dangerous-direct-browser-access': 'true',
    }),
    getEndpoint: () => 'https://api.anthropic.com/v1/messages',
    parseResponse: (data) => data?.content?.[0]?.text ?? '',
  },
  openai: openAiCompatibleAdapter,
  gemini: {
    formatRequest: (messages, settings) => ({
      contents: messages.map((message) => ({
        role: message.role === 'assistant' ? 'model' : 'user',
        parts: [{ text: message.content }],
      })),
      systemInstruction: settings.systemPrompt
        ? { parts: [{ text: settings.systemPrompt }] }
        : undefined,
      generationConfig: {
        temperature: settings.temperature,
        maxOutputTokens: settings.maxTokens,
      },
    }),
    getHeaders: () => ({
      'Content-Type': 'application/json',
    }),
    getEndpoint: (model, _, apiKey) =>
      `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`,
    parseResponse: (data) =>
      data?.candidates?.[0]?.content?.parts?.[0]?.text ?? '',
  },
  grok: openAiCompatibleAdapter,
  mistral: openAiCompatibleAdapter,
  cohere: {
    formatRequest: (messages, settings) => {
      const lastMessage = messages[messages.length - 1]
      const history = messages.slice(0, -1)
      return {
        model: settings.model,
        message: lastMessage?.content ?? '',
        chat_history: history.map((message) => ({
          role: message.role === 'user' ? 'USER' : 'CHATBOT',
          message: message.content,
        })),
        preamble: settings.systemPrompt,
        temperature: settings.temperature,
        max_tokens: settings.maxTokens,
      }
    },
    getHeaders: (apiKey) => ({
      'Content-Type': 'application/json',
      Authorization: `Bearer ${apiKey}`,
    }),
    getEndpoint: () => 'https://api.cohere.ai/v1/chat',
    parseResponse: (data) => data?.text ?? '',
  },
}


const MessageBubble = ({ message, onError }) => {
  const isUser = message.role === 'user'
  const provider = message.provider ? PROVIDERS[message.provider] : null
  const modelLabel =
    provider && message.model ? getModelName(provider.id, message.model) : ''
  const timestamp = formatTime(message.timestamp)
  const attachments = Array.isArray(message.attachments)
    ? message.attachments
    : []
  const isSoloCodeBlock =
    !isUser && isSoloMarkdownCodeBlock(message.content)
  const markdownComponents = createMarkdownComponents(isSoloCodeBlock)
  const markdownContent = !isUser
    ? normalizeMathDelimiters(message.content)
    : message.content
  const [copiedResponse, setCopiedResponse] = useState(false)

  const handleCopyResponse = async () => {
    if (!message.content) {
      return
    }
    const normalizedContent = normalizeMathForClipboard(message.content)
    if (navigator.clipboard && window.isSecureContext && window.ClipboardItem) {
      try {
        const html = buildExportHtml(normalizedContent)
        const item = new ClipboardItem({
          'text/plain': new Blob([normalizedContent], { type: 'text/plain' }),
          'text/html': new Blob([html], { type: 'text/html' }),
        })
        await navigator.clipboard.write([item])
        setCopiedResponse(true)
        setTimeout(() => setCopiedResponse(false), 1200)
        return
      } catch (error) {
        // Fall back to plain text copy.
      }
    }

    const success = await copyToClipboard(normalizedContent)
    if (success) {
      setCopiedResponse(true)
      setTimeout(() => setCopiedResponse(false), 1200)
    } else {
      onError?.('Unable to copy response.')
    }
  }

  const handleExportPdf = async () => {
    if (!message.content) {
      return
    }
    const title = getResponseTitle(message.content)
    const timestamp = formatExportTimestamp()
    const filename = `${title} - ${timestamp}`
    try {
      await exportMarkdownAsPdf(message.content, filename)
    } catch (error) {
      console.error('PDF export failed', error)
      const detail =
        error && typeof error === 'object' && 'message' in error
          ? `Failed to export PDF: ${error.message}`
          : 'Failed to export PDF.'
      onError?.(detail)
    }
  }

  return (
    <div className={`flex ${isUser ? 'justify-end' : 'justify-start'}`}>
      <div
        className={`${
          isUser
            ? 'max-w-[70%] rounded-[18px] rounded-br-[6px] bg-[#007AFF] px-4 py-3 text-sm text-white shadow-[0_2px_12px_rgba(15,23,42,0.08)]'
            : 'w-full px-2 py-2 text-sm text-slate-900 dark:text-[#ffffff]'
        }`}
      >
        {!isUser && provider ? (
          <div className="mb-1 flex items-center gap-2 text-[11px] font-semibold text-slate-600 dark:text-[#afafaf]">
            <span
              className="h-2.5 w-2.5 rounded-full"
              style={{ backgroundColor: provider.color }}
            />
            <span>{provider.name}</span>
            {modelLabel ? (
              <span className="text-slate-500 dark:text-[#afafaf]">
                - {modelLabel}
              </span>
            ) : null}
          </div>
        ) : null}
        {isUser && attachments.length ? (
          <div className="mb-2 space-y-1 rounded-xl border border-white/30 bg-white/15 px-3 py-2 text-xs text-white/85 dark:border-[#fff3] dark:bg-[#303030] dark:text-[#f3f3f3]">
            {attachments.map((attachment) => (
              <div key={attachment.id} className="flex items-center gap-2">
                <Paperclip className="h-3 w-3" />
                <span className="truncate">{attachment.name}</span>
                <span className="text-[10px] opacity-70">
                  {attachment.kind.toUpperCase()}
                </span>
              </div>
            ))}
          </div>
        ) : null}
        <div className="leading-relaxed">
          {isUser ? (
            <div className="whitespace-pre-wrap">
              {message.displayContent ?? message.content}
            </div>
          ) : (
            <ReactMarkdown
              remarkPlugins={[remarkGfm, remarkMath]}
              rehypePlugins={[rehypeKatex]}
              components={markdownComponents}
            >
              {markdownContent}
            </ReactMarkdown>
          )}
        </div>
        <div
          className={`mt-2 flex items-center ${
            isUser ? 'justify-end' : 'justify-between'
          }`}
        >
          {!isUser ? (
            <div className="flex items-center gap-2 text-slate-500 dark:text-[#afafaf]">
              <button
                type="button"
                onClick={handleCopyResponse}
                aria-label="Copy response"
                className="rounded-full border border-slate-200 p-1 transition hover:border-slate-300 hover:text-slate-700 dark:border-[#ffffff26] dark:hover:border-[#fff3] dark:hover:text-[#f3f3f3]"
                title="Copy response"
              >
                <Copy
                  className={`h-3.5 w-3.5 ${
                    copiedResponse ? 'text-emerald-400' : ''
                  }`}
                />
              </button>
              <button
                type="button"
                onClick={handleExportPdf}
                aria-label="Export response as PDF"
                className="rounded-full border border-slate-200 p-1 transition hover:border-slate-300 hover:text-slate-700 dark:border-[#ffffff26] dark:hover:border-[#fff3] dark:hover:text-[#f3f3f3]"
                title="Export PDF"
              >
                <Download className="h-3.5 w-3.5" />
              </button>
            </div>
          ) : (
            <span />
          )}
          <div
            className={`text-[11px] ${
              isUser
                ? 'text-white/70 text-right'
                : 'text-slate-500 dark:text-[#afafaf]'
            }`}
          >
            {timestamp}
          </div>
        </div>
      </div>
    </div>
  )
}

const TypingIndicator = ({ providerId }) => {
  const provider = PROVIDERS[providerId]
  return (
    <div className="flex justify-start">
      <div className="w-full px-2 py-2 text-sm text-slate-600 dark:text-[#afafaf]">
        <div className="mb-1 flex items-center gap-2 text-[11px] font-semibold">
          <span
            className="h-2.5 w-2.5 rounded-full"
            style={{ backgroundColor: provider.color }}
          />
          <span>{provider.name}</span>
          <span className="text-slate-500 dark:text-[#afafaf]">is typing</span>
        </div>
        <div className="flex items-center gap-1">
          <span
            className="h-2 w-2 rounded-full bg-slate-400 animate-bounce"
            style={{ animationDelay: '0ms' }}
          />
          <span
            className="h-2 w-2 rounded-full bg-slate-400 animate-bounce"
            style={{ animationDelay: '120ms' }}
          />
          <span
            className="h-2 w-2 rounded-full bg-slate-400 animate-bounce"
            style={{ animationDelay: '240ms' }}
          />
        </div>
      </div>
    </div>
  )
}

const EmptyState = ({ providerId }) => {
  const provider = PROVIDERS[providerId]
  return (
    <div className="flex h-full flex-col items-center justify-center gap-4 py-12 text-center">
      <div className="flex h-14 w-14 items-center justify-center rounded-2xl bg-slate-900 text-white shadow-lg dark:bg-[#414141] dark:text-[#ffffff]">
        <MessageSquare className="h-6 w-6" />
      </div>
      <div className="max-w-md space-y-2">
        <p className="text-lg font-semibold text-slate-900 dark:text-[#ffffff]">
          Start a new conversation
        </p>
        <p className="text-sm text-slate-500 dark:text-[#afafaf]">
          Choose a provider, add your API key, and send your first message to
          get going.
        </p>
      </div>
      <div className="flex flex-wrap items-center justify-center gap-2">
        {PROVIDER_ORDER.map((providerKey) => {
          const item = PROVIDERS[providerKey]
          return (
            <div
              key={providerKey}
              className="flex items-center gap-2 rounded-full border border-slate-200 bg-white/90 px-3 py-1 text-xs text-slate-600 shadow-sm dark:border-[#ffffff26] dark:bg-[#303030]/80 dark:text-[#cdcdcd]"
            >
              <span
                className="h-2.5 w-2.5 rounded-full"
                style={{ backgroundColor: item.color }}
              />
              <span>{item.name}</span>
            </div>
          )
        })}
      </div>
      <div className="text-xs text-slate-500 dark:text-[#afafaf]">
        Active provider: {provider.name}
      </div>
    </div>
  )
}

const ErrorToast = ({ message, onDismiss }) => {
  if (!message) {
    return null
  }
  return (
    <div className="fixed right-4 top-4 z-50 w-[320px] rounded-2xl border border-red-100 bg-white/95 p-4 shadow-[0_20px_60px_rgba(15,23,42,0.15)] backdrop-blur dark:border-red-900/50 dark:bg-[#181818]/90">
      <div className="flex items-start gap-3">
        <div className="flex h-9 w-9 items-center justify-center rounded-full bg-red-50 text-red-600 dark:bg-red-950/40 dark:text-red-300">
          <AlertTriangle className="h-4 w-4" />
        </div>
        <div className="flex-1 text-sm text-slate-700 dark:text-[#f3f3f3]">
          {message}
        </div>
        <button
          type="button"
          onClick={onDismiss}
          className="text-slate-400 transition hover:text-slate-700 dark:text-[#afafaf] dark:hover:text-[#f3f3f3]"
          aria-label="Dismiss error"
        >
          <X className="h-4 w-4" />
        </button>
      </div>
    </div>
  )
}

function App() {
  const [activeProvider, setActiveProvider] = useState(() => {
    const stored = loadStored('llm-chat-provider', 'anthropic')
    return PROVIDERS[stored] ? stored : 'anthropic'
  })
  const [activeModel, setActiveModel] = useState(() =>
    loadStored('llm-chat-model', PROVIDERS.anthropic.models[0].id),
  )
  const [apiKeys, setApiKeys] = useState(() => {
    const keys = {}
    PROVIDER_ORDER.forEach((providerId) => {
      keys[providerId] = loadStored(`llm-chat-key-${providerId}`, '')
    })
    return keys
  })
  const [temperature, setTemperature] = useState(() =>
    Number(loadStored('llm-chat-temperature', 1)),
  )
  const [maxTokens, setMaxTokens] = useState(() =>
    Number(loadStored('llm-chat-max-tokens', 1024)),
  )
  const [systemPrompt, setSystemPrompt] = useState(() =>
    loadStored('llm-chat-system-prompt', DEFAULT_SYSTEM_PROMPT),
  )
  const [reasoningEffort, setReasoningEffort] = useState(() =>
    loadStored('llm-chat-reasoning-effort', 'high'),
  )
  const [verbosity, setVerbosity] = useState(() =>
    loadStored('llm-chat-verbosity', 'high'),
  )
  const [reasoningPreviewEnabled, setReasoningPreviewEnabled] = useState(() =>
    loadStored('llm-chat-reasoning-preview', true),
  )
  const [sidebarOpen, setSidebarOpen] = useState(() =>
    loadStored('llm-chat-sidebar-open', true),
  )
  const [darkMode, setDarkMode] = useState(
    () => loadStored('llm-chat-theme', 'light') === 'dark',
  )
  const [attachments, setAttachments] = useState([])
  const [chats, setChats] = useState(() => {
    const storedChats = loadStored('llm-chat-chats', null)
    if (Array.isArray(storedChats) && storedChats.length) {
      return storedChats.map(normalizeChat)
    }
    const legacyMessages = loadStored('llm-chat-messages', [])
    return [createChat(Array.isArray(legacyMessages) ? legacyMessages : [])]
  })
  const [activeChatId, setActiveChatId] = useState(() =>
    loadStored('llm-chat-active-chat', null),
  )
  const [inputValue, setInputValue] = useState('')
  const [settingsOpen, setSettingsOpen] = useState(false)
  const [isLoading, setIsLoading] = useState(false)
  const [streamingMessageId, setStreamingMessageId] = useState(null)
  const [error, setError] = useState(null)
  const [showApiKeys, setShowApiKeys] = useState(() => {
    const state = {}
    PROVIDER_ORDER.forEach((providerId) => {
      state[providerId] = false
    })
    return state
  })

  const inputRef = useRef(null)
  const fileInputRef = useRef(null)
  const listRef = useRef(null)
  const listOuterRef = useRef(null)
  const listContainerRef = useRef(null)
  const sizeMapRef = useRef(new Map())

  const provider = PROVIDERS[activeProvider] ?? PROVIDERS.anthropic
  const keyWarning = getKeyWarning(activeProvider, apiKeys[activeProvider])
  const showGpt52Controls =
    activeProvider === 'openai' && activeModel.startsWith('gpt-5.2')
  const showGpt5Controls =
    activeProvider === 'openai' && activeModel.startsWith('gpt-5')
  const normalizedReasoningEffort = normalizeReasoningEffort(
    activeModel,
    reasoningEffort,
  )
  const temperatureDisabled =
    showGpt52Controls && normalizedReasoningEffort !== 'none'
  const canSend = Boolean(inputValue.trim() || attachments.length)
  const showTypingIndicator = isLoading && !streamingMessageId
  const orderedChats = [...chats].sort((a, b) => {
    const aTime = new Date(a.updatedAt || a.createdAt || 0).getTime()
    const bTime = new Date(b.updatedAt || b.createdAt || 0).getTime()
    return bTime - aTime
  })
  const activeChat =
    chats.find((chat) => chat.id === activeChatId) || chats[0]
  const messages = activeChat?.messages ?? []
  const virtualItems = useMemo(
    () =>
      showTypingIndicator
        ? [...messages, { id: 'typing-indicator', type: 'typing' }]
        : messages,
    [messages, showTypingIndicator],
  )
  const setSize = useCallback((index, size) => {
    const stored = sizeMapRef.current.get(index)
    if (stored !== size) {
      sizeMapRef.current.set(index, size)
      listRef.current?.resetAfterIndex(index)
    }
  }, [])

  const getSize = useCallback(
    (index) => sizeMapRef.current.get(index) ?? VIRTUAL_ESTIMATE,
    [],
  )
  const [listSize, setListSize] = useState({ height: 0, width: 0 })
  const [isAtBottom, setIsAtBottom] = useState(true)
  const showVirtualList = listSize.height > 0 && listSize.width > 0
  const useVirtualList = showVirtualList && listSize.width >= 640
  const listData = useMemo(
    () => ({
      items: virtualItems,
      onError: setError,
      providerId: activeProvider,
      setSize,
    }),
    [virtualItems, setError, activeProvider, setSize],
  )

  useEffect(() => {
    saveStored('llm-chat-provider', activeProvider)
  }, [activeProvider])

  useEffect(() => {
    saveStored('llm-chat-model', activeModel)
  }, [activeModel])

  useEffect(() => {
    saveStored('llm-chat-temperature', temperature)
  }, [temperature])

  useEffect(() => {
    saveStored('llm-chat-max-tokens', maxTokens)
  }, [maxTokens])

  useEffect(() => {
    saveStored('llm-chat-system-prompt', systemPrompt)
  }, [systemPrompt])

  useEffect(() => {
    saveStored('llm-chat-reasoning-effort', reasoningEffort)
  }, [reasoningEffort])

  useEffect(() => {
    saveStored('llm-chat-verbosity', verbosity)
  }, [verbosity])

  useEffect(() => {
    saveStored('llm-chat-reasoning-preview', reasoningPreviewEnabled)
  }, [reasoningPreviewEnabled])

  useEffect(() => {
    saveStored('llm-chat-sidebar-open', sidebarOpen)
  }, [sidebarOpen])

  useEffect(() => {
    const root = document.documentElement
    if (darkMode) {
      root.classList.add('dark')
    } else {
      root.classList.remove('dark')
    }
    saveStored('llm-chat-theme', darkMode ? 'dark' : 'light')
  }, [darkMode])

  useEffect(() => {
    saveStored('llm-chat-chats', chats)
  }, [chats])

  useEffect(() => {
    if (activeChatId) {
      saveStored('llm-chat-active-chat', activeChatId)
    }
  }, [activeChatId])

  useEffect(() => {
    PROVIDER_ORDER.forEach((providerId) => {
      saveStored(`llm-chat-key-${providerId}`, apiKeys[providerId] ?? '')
    })
  }, [apiKeys])

  useEffect(() => {
    const providerSettings = PROVIDERS[activeProvider]
    if (!providerSettings) {
      return
    }
    if (!providerSettings.models.some((model) => model.id === activeModel)) {
      setActiveModel(providerSettings.models[0].id)
    }
    setTemperature((current) =>
      Number.isFinite(current)
        ? clamp(
            current,
            providerSettings.tempRange.min,
            providerSettings.tempRange.max,
          )
        : providerSettings.tempRange.default,
    )
    setMaxTokens((current) =>
      Number.isFinite(current)
        ? clamp(
            current,
            providerSettings.maxTokensRange.min,
            providerSettings.maxTokensRange.max,
          )
        : providerSettings.maxTokensRange.default,
    )
  }, [activeProvider, activeModel])

  useEffect(() => {
    if (!activeChat && chats.length) {
      setActiveChatId(chats[0].id)
    }
  }, [activeChat, chats])

  useLayoutEffect(() => {
    const container = listContainerRef.current
    if (!container) {
      return
    }
    const update = () => {
      setListSize({
        height: container.clientHeight,
        width: container.clientWidth,
      })
    }
    update()
    if (typeof ResizeObserver === 'undefined') {
      return
    }
    const observer = new ResizeObserver(update)
    observer.observe(container)
    return () => observer.disconnect()
  }, [])

  useEffect(() => {
    if (!virtualItems.length || !isAtBottom) {
      return
    }
    if (useVirtualList && listRef.current) {
      listRef.current.scrollToItem(virtualItems.length - 1, 'end')
      return
    }
    if (listOuterRef.current) {
      listOuterRef.current.scrollTop = listOuterRef.current.scrollHeight
    }
  }, [virtualItems.length, isAtBottom, useVirtualList])

  useEffect(() => {
    sizeMapRef.current = new Map()
    listRef.current?.resetAfterIndex(0, true)
    setIsAtBottom(true)
  }, [activeChatId])

  useEffect(() => {
    if (!error) {
      return
    }
    const timeout = setTimeout(() => {
      setError(null)
    }, 5000)
    return () => clearTimeout(timeout)
  }, [error])

  useEffect(() => {
    const handleKeyDown = (event) => {
      if (
        (event.metaKey || event.ctrlKey) &&
        event.key.toLowerCase() === 'k'
      ) {
        event.preventDefault()
        setSettingsOpen(true)
      }
      if (event.key === 'Escape') {
        setSettingsOpen(false)
      }
    }
    window.addEventListener('keydown', handleKeyDown)
    return () => window.removeEventListener('keydown', handleKeyDown)
  }, [])

  useEffect(() => {
    const normalized = normalizeReasoningEffort(activeModel, reasoningEffort)
    if (normalized !== reasoningEffort) {
      setReasoningEffort(normalized)
    }
  }, [activeModel, reasoningEffort])

  const handleSend = async () => {
    if (isLoading) {
      return
    }
    const trimmed = inputValue.trim()
    if (!trimmed && attachments.length === 0) {
      return
    }
    const currentChat = activeChat || createChat([])
    if (!activeChat) {
      setChats((prev) => [currentChat, ...prev])
      setActiveChatId(currentChat.id)
    }
    const apiKey = apiKeys[activeProvider]
    if (!apiKey) {
      setError(`Missing API key for ${provider.name}. Please add it first.`)
      return
    }

    const attachmentContext = attachments.length
      ? `\n\nAttached context:\n${attachments
          .map(formatAttachmentForPrompt)
          .join('\n\n')}`
      : ''
    const displayContent = trimmed || 'Shared attachments'
    const userMessage = {
      id: generateId(),
      role: 'user',
      content: `${trimmed}${attachmentContext}`.trim(),
      displayContent,
      timestamp: new Date().toISOString(),
      attachments: attachments.map((item) => ({
        id: item.id,
        name: item.name,
        kind: item.kind,
        size: item.size,
      })),
    }
    const conversation = [...(currentChat.messages || []), userMessage]
    setChats((prev) =>
      prev.map((chat) => {
        if (chat.id !== currentChat.id) {
          return chat
        }
        const nextTitle =
          conversation.length === 0
            ? 'New chat'
            : chat.title === 'New chat'
              ? deriveChatTitle(conversation)
              : chat.title
        return {
          ...chat,
          messages: conversation,
          updatedAt: new Date().toISOString(),
          title: nextTitle,
        }
      }),
    )
    setInputValue('')
    setAttachments([])
    if (inputRef.current) {
      inputRef.current.style.height = 'auto'
    }
    setIsLoading(true)

    try {
      const adapter = ADAPTERS[activeProvider]
      const settings = {
        model: activeModel,
        temperature,
        maxTokens,
        systemPrompt,
        reasoningEffort,
        verbosity,
        reasoningPreviewEnabled: showGpt5Controls && reasoningPreviewEnabled,
      }
      const body = adapter.formatRequest(conversation, settings)
      const endpoint = adapter.getEndpoint(
        activeModel,
        provider.endpoint,
        apiKey,
      )
      const shouldStream = STREAMING_PROVIDERS.has(activeProvider)
      const requestBody = shouldStream ? { ...body, stream: true } : body
      const headers = { ...adapter.getHeaders(apiKey) }
      if (shouldStream) {
        headers.Accept = 'text/event-stream'
      }

      const response = await fetch(endpoint, {
        method: 'POST',
        headers,
        body: JSON.stringify(requestBody),
      })

      if (!response.ok) {
        let errorMessage = `${provider.name} request failed.`
        try {
          const errorData = await response.json()
          const providerMessage =
            errorData?.error?.message ||
            errorData?.message ||
            errorData?.error?.details?.[0]?.message
          if (providerMessage) {
            errorMessage = providerMessage
          }
        } catch (parseError) {
          // Ignore parse errors and use fallback message.
        }

        if (response.status === 401 || response.status === 403) {
          errorMessage = `Invalid API key for ${provider.name}. Please check your settings.`
        } else if (response.status === 429) {
          errorMessage = 'Rate limited. Please wait a moment.'
        }
        throw new Error(errorMessage)
      }

      if (shouldStream && response.body) {
        const assistantId = generateId()
        const assistantMessage = {
          id: assistantId,
          role: 'assistant',
          content: '',
          timestamp: new Date().toISOString(),
          provider: activeProvider,
          model: activeModel,
          reasoningPreview: '',
        }
        setChats((prev) =>
          prev.map((chat) =>
            chat.id === currentChat.id
              ? {
                  ...chat,
                  messages: [...chat.messages, assistantMessage],
                  updatedAt: new Date().toISOString(),
                }
              : chat,
          ),
        )
        setStreamingMessageId(assistantId)

        const reader = response.body.getReader()
        const decoder = new TextDecoder()
        let buffer = ''
        let eventName = ''
        let dataLines = []
        let fullText = ''

        const updateAssistantContent = (content, preview = '') => {
          setChats((prev) =>
            prev.map((chat) =>
              chat.id === currentChat.id
                ? {
                    ...chat,
                    messages: chat.messages.map((message) =>
                      message.id === assistantId
                        ? { ...message, content, reasoningPreview: preview }
                        : message,
                    ),
                    updatedAt: new Date().toISOString(),
                  }
                : chat,
            ),
          )
        }

        const handleStreamEvent = (dataText, name) => {
          if (!dataText) {
            return false
          }
          if (dataText === '[DONE]') {
            return true
          }
          let payload = null
          try {
            payload = JSON.parse(dataText)
          } catch (error) {
            return false
          }
          const delta = extractStreamDelta(name, payload)
          if (delta) {
            fullText += delta
            updateAssistantContent(fullText)
          }
          return false
        }

        let done = false
        while (!done) {
          const { value, done: streamDone } = await reader.read()
          if (streamDone) {
            break
          }
          buffer += decoder.decode(value, { stream: true })
          const lines = buffer.split(/\r?\n/)
          buffer = lines.pop() ?? ''
          for (const line of lines) {
            if (!line) {
              const dataText = dataLines.join('\n')
              done = handleStreamEvent(dataText, eventName)
              eventName = ''
              dataLines = []
              if (done) {
                break
              }
              continue
            }
            if (line.startsWith(':')) {
              continue
            }
            if (line.startsWith('event:')) {
              eventName = line.slice(6).trim()
              continue
            }
            if (line.startsWith('data:')) {
              dataLines.push(line.slice(5).trim())
              continue
            }
          }
        }

        if (dataLines.length) {
          handleStreamEvent(dataLines.join('\n'), eventName)
        }

        if (!fullText.trim()) {
          throw new Error(`Empty response from ${provider.name}.`)
        }

        const previewResult =
          showGpt5Controls && reasoningPreviewEnabled
            ? extractReasoningPreview(fullText)
            : { cleaned: fullText, preview: '' }
        updateAssistantContent(previewResult.cleaned, previewResult.preview)
        return
      }

      const data = await response.json()
      if (
        data?.status === 'incomplete' &&
        data?.incomplete_details?.reason === 'max_output_tokens'
      ) {
        throw new Error(
          'Response hit max output tokens. Increase Max Tokens or lower reasoning effort/verbosity.',
        )
      }
      const content = adapter.parseResponse(data)
      if (!content) {
        throw new Error(`Empty response from ${provider.name}.`)
      }

      const previewResult =
        showGpt5Controls && reasoningPreviewEnabled
          ? extractReasoningPreview(content)
          : { cleaned: content, preview: '' }
      const assistantMessage = {
        id: generateId(),
        role: 'assistant',
        content: previewResult.cleaned,
        timestamp: new Date().toISOString(),
        provider: activeProvider,
        model: activeModel,
        reasoningPreview: previewResult.preview,
      }
      const nextConversation = [...conversation, assistantMessage]
      setChats((prev) =>
        prev.map((chat) =>
          chat.id === currentChat.id
            ? {
                ...chat,
                messages: nextConversation,
                updatedAt: new Date().toISOString(),
              }
            : chat,
        ),
      )
    } catch (requestError) {
      setError(requestError?.message || 'Request failed. Please try again.')
    } finally {
      setIsLoading(false)
      setStreamingMessageId(null)
    }
  }

  const handleInputChange = (event) => {
    setInputValue(event.target.value)
    if (inputRef.current) {
      inputRef.current.style.height = 'auto'
      inputRef.current.style.height = `${Math.min(
        inputRef.current.scrollHeight,
        160,
      )}px`
    }
  }

  const handleInputKeyDown = (event) => {
    if (event.key === 'Enter' && !event.shiftKey) {
      event.preventDefault()
      handleSend()
    }
  }

  const handleAttachClick = () => {
    fileInputRef.current?.click()
  }

  const handleFilesSelected = async (event) => {
    const files = Array.from(event.target.files || [])
    event.target.value = ''
    if (!files.length) {
      return
    }
    const nextAttachments = []
    for (const file of files) {
      try {
        const { text, kind } = await readAttachmentText(file)
        nextAttachments.push({
          id: generateId(),
          name: file.name,
          size: file.size,
          type: file.type,
          kind,
          text,
        })
      } catch (error) {
        setError(`Failed to read ${file.name}.`)
      }
    }
    if (nextAttachments.length) {
      setAttachments((prev) => [...prev, ...nextAttachments])
    }
  }

  const handleRemoveAttachment = (attachmentId) => {
    setAttachments((prev) =>
      prev.filter((attachment) => attachment.id !== attachmentId),
    )
  }

  const handleNewChat = () => {
    const newChat = createChat([])
    setChats((prev) => [newChat, ...prev])
    setActiveChatId(newChat.id)
    setInputValue('')
    setAttachments([])
    if (inputRef.current) {
      inputRef.current.style.height = 'auto'
    }
  }

  const handleClearChat = () => {
    if (!activeChat) {
      return
    }
    setChats((prev) =>
      prev.map((chat) =>
        chat.id === activeChat.id
          ? {
              ...chat,
              messages: [],
              updatedAt: new Date().toISOString(),
              title: 'New chat',
            }
          : chat,
      ),
    )
  }

  const handleExport = () => {
    if (!activeChat || !messages.length) {
      return
    }
    const payload = {
      id: activeChat.id,
      title: activeChat.title,
      messages,
    }
    const blob = new Blob([JSON.stringify(payload, null, 2)], {
      type: 'application/json',
    })
    const url = URL.createObjectURL(blob)
    const link = document.createElement('a')
    link.href = url
    link.download = 'conversation.json'
    document.body.appendChild(link)
    link.click()
    link.remove()
    URL.revokeObjectURL(url)
  }

  return (
    <div className="min-h-[100svh] w-full overflow-x-hidden bg-[#f7f7f9] text-slate-900 dark:bg-[#303030] dark:text-[#ffffff]">
      <div className="relative mx-auto flex min-h-[100svh] w-full max-w-6xl flex-col px-4 py-6 font-display">
        <div className="pointer-events-none absolute inset-0">
          <div className="absolute right-8 top-12 h-48 w-48 rounded-full bg-white/60 blur-2xl dark:bg-white/5" />
          <div className="absolute bottom-12 left-12 h-56 w-56 rounded-full bg-slate-200/40 blur-3xl dark:bg-white/6" />
        </div>

        <header className="relative z-10 flex flex-wrap items-center justify-between gap-4 rounded-3xl border border-slate-200 bg-white/85 p-4 shadow-[0_16px_40px_rgba(15,23,42,0.08)] backdrop-blur dark:border-[#ffffff26] dark:bg-[#303030]/80">
          <div className="flex items-center gap-3">
            <div className="flex h-12 w-12 items-center justify-center rounded-2xl bg-slate-900 text-white shadow-md dark:bg-[#414141] dark:text-[#ffffff]">
              <MessageSquare className="h-5 w-5" />
            </div>
            <div>
              <div className="text-lg font-semibold">Multi-LLM Chat</div>
              <div className="text-xs text-slate-500 dark:text-[#afafaf]">
                Switch providers mid-thread
              </div>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-3">
            <div className="flex items-center gap-2 rounded-full border border-slate-200 bg-white/70 px-3 py-2 text-sm shadow-sm dark:border-[#ffffff26] dark:bg-[#303030]/70">
              <span
                className="h-2.5 w-2.5 rounded-full"
                style={{ backgroundColor: provider.color }}
              />
              <select
                className="bg-transparent text-sm font-medium text-slate-700 outline-none dark:bg-[#303030] dark:text-[#ffffff] [&>option]:bg-white [&>option]:text-slate-900 dark:[&>option]:bg-[#303030] dark:[&>option]:text-[#ffffff]"
                value={activeProvider}
                onChange={(event) => setActiveProvider(event.target.value)}
              >
                {PROVIDER_ORDER.map((providerId) => (
                  <option key={providerId} value={providerId}>
                    {PROVIDERS[providerId].name}
                  </option>
                ))}
              </select>
            </div>

            <div className="flex items-center gap-2 rounded-full border border-slate-200 bg-white/70 px-3 py-2 text-sm shadow-sm dark:border-[#ffffff26] dark:bg-[#303030]/70">
              <span className="text-xs text-slate-500 dark:text-[#afafaf]">Model</span>
              <select
                className="bg-transparent text-sm font-medium text-slate-700 outline-none dark:bg-[#303030] dark:text-[#ffffff] [&>option]:bg-white [&>option]:text-slate-900 dark:[&>option]:bg-[#303030] dark:[&>option]:text-[#ffffff]"
                value={activeModel}
                onChange={(event) => setActiveModel(event.target.value)}
              >
                {provider.models.map((model) => (
                  <option key={model.id} value={model.id}>
                    {model.name}
                  </option>
                ))}
              </select>
            </div>

            <button
              type="button"
              role="switch"
              aria-checked={darkMode}
              aria-label="Toggle dark mode"
              onClick={() => setDarkMode((current) => !current)}
              className="relative inline-flex h-9 w-16 items-center rounded-full border border-slate-200 bg-white/70 px-1 shadow-sm transition hover:border-slate-300 dark:border-[#ffffff26] dark:bg-[#303030]/70 dark:hover:border-[#fff3]"
            >
              <span
                className={`flex h-7 w-7 items-center justify-center rounded-full text-slate-600 shadow-sm transition ${
                  darkMode
                    ? 'translate-x-7 bg-slate-100 text-slate-900'
                    : 'translate-x-0 bg-white'
                }`}
              >
                {darkMode ? (
                  <Moon className="h-3 w-3" />
                ) : (
                  <Sun className="h-3 w-3" />
                )}
              </span>
            </button>

            <button
              type="button"
              onClick={() => setSidebarOpen((current) => !current)}
              aria-expanded={sidebarOpen}
              aria-controls="chat-sidebar"
              aria-label={sidebarOpen ? 'Close sidebar' : 'Open sidebar'}
              className="flex h-9 w-9 items-center justify-center rounded-full border border-slate-200 bg-white/70 text-slate-700 shadow-sm transition hover:border-slate-300 hover:bg-white dark:border-[#ffffff26] dark:bg-[#303030]/70 dark:text-[#f3f3f3] dark:hover:border-[#fff3] dark:hover:bg-[#303030]"
            >
              <PanelLeft className="h-4 w-4" />
            </button>

            <button
              type="button"
              onClick={() => setSettingsOpen(true)}
              className="flex items-center gap-2 rounded-full border border-slate-200 bg-white/70 px-3 py-2 text-sm font-medium text-slate-700 shadow-sm transition hover:border-slate-300 hover:bg-white dark:border-[#ffffff26] dark:bg-[#303030]/70 dark:text-[#f3f3f3] dark:hover:border-[#fff3] dark:hover:bg-[#303030]"
            >
              <Settings className="h-4 w-4" />
              Settings
            </button>
          </div>
        </header>

        <main className="relative z-10 mt-6 flex w-full min-w-0 flex-1 flex-col gap-4 min-h-0 lg:flex-row">
          {sidebarOpen ? (
            <aside
              id="chat-sidebar"
              className="flex w-full min-w-0 flex-col rounded-3xl border border-slate-200 bg-white/85 p-4 shadow-[0_16px_40px_rgba(15,23,42,0.08)] backdrop-blur dark:border-[#ffffff26] dark:bg-[#303030]/80 lg:w-72"
            >
              <div className="flex items-center justify-between">
                <div className="text-xs font-semibold uppercase tracking-wide text-slate-500 dark:text-[#afafaf]">
                  Chats
                </div>
                <button
                  type="button"
                  onClick={handleNewChat}
                  className="rounded-full border border-slate-200 px-3 py-1 text-xs font-semibold text-slate-600 transition hover:border-slate-300 hover:text-slate-900 dark:border-[#ffffff26] dark:text-[#cdcdcd] dark:hover:border-[#fff3] dark:hover:text-[#ffffff]"
                >
                  New chat
                </button>
              </div>
              <div className="mt-4 flex-1 space-y-2 overflow-y-auto pr-1">
                {orderedChats.map((chat) => {
                  const isActive = chat.id === activeChat?.id
                  return (
                    <button
                      key={chat.id}
                      type="button"
                      onClick={() => setActiveChatId(chat.id)}
                      className={`w-full rounded-2xl border px-3 py-2 text-left transition ${
                        isActive
                          ? 'border-slate-900 bg-slate-900 text-white dark:border-[#fff3] dark:bg-[#414141] dark:text-[#ffffff]'
                          : 'border-slate-200 bg-white text-slate-700 hover:border-slate-300 dark:border-[#ffffff26] dark:bg-[#303030] dark:text-[#f3f3f3] dark:hover:border-[#fff3]'
                      }`}
                    >
                      <div className="flex items-center justify-between gap-2">
                        <span
                          className={`text-sm font-semibold ${
                            isActive
                              ? 'text-white dark:text-[#ffffff]'
                              : 'text-slate-700 dark:text-[#f3f3f3]'
                          }`}
                        >
                          {chat.title}
                        </span>
                        <span
                          className={`text-[10px] ${
                            isActive
                              ? 'text-white/70 dark:text-[#cdcdcd]'
                              : 'text-slate-400 dark:text-[#afafaf]'
                          }`}
                        >
                          {formatTime(chat.updatedAt || chat.createdAt)}
                        </span>
                      </div>
                      <div
                        className={`mt-1 text-xs ${
                          isActive
                            ? 'text-white/80 dark:text-[#afafaf]'
                            : 'text-slate-500 dark:text-[#afafaf]'
                        }`}
                      >
                        {getChatPreview(chat.messages)}
                      </div>
                    </button>
                  )
                })}
              </div>
            </aside>
          ) : null}

          <div className="flex min-w-0 flex-1 flex-col gap-4 min-h-0">
            <section className="flex min-h-0 flex-1 flex-col overflow-hidden rounded-3xl border border-slate-200 bg-white/85 shadow-[0_16px_40px_rgba(15,23,42,0.08)] backdrop-blur dark:border-[#ffffff26] dark:bg-[#303030]/80">
              <div ref={listContainerRef} className="flex-1 min-h-0 w-full">
                {messages.length === 0 ? (
                  <div className="flex h-full flex-col">
                    <EmptyState providerId={activeProvider} />
                  </div>
                ) : useVirtualList ? (
                  <List
                    ref={listRef}
                    outerRef={listOuterRef}
                    height={listSize.height}
                    width={listSize.width}
                    itemCount={virtualItems.length}
                    itemSize={getSize}
                    itemKey={(index, data) =>
                      data.items[index]?.id || `row-${index}`
                    }
                    onScroll={() => {
                      const outer = listOuterRef.current
                      if (!outer) {
                        return
                      }
                      const distance =
                        outer.scrollHeight - outer.scrollTop - outer.clientHeight
                      setIsAtBottom(distance < 80)
                    }}
                    itemData={listData}
                    outerElementType={ListOuter}
                  >
                    {VirtualRow}
                  </List>
                ) : (
                  <div
                    ref={listOuterRef}
                    onScroll={() => {
                      const outer = listOuterRef.current
                      if (!outer) {
                        return
                      }
                      const distance =
                        outer.scrollHeight - outer.scrollTop - outer.clientHeight
                      setIsAtBottom(distance < 80)
                    }}
                    className="h-full min-h-0 overflow-y-auto px-6 py-5"
                  >
                    {virtualItems.map((item) =>
                      item?.type === 'typing' ? (
                        <TypingIndicator
                          key={item.id}
                          providerId={activeProvider}
                        />
                      ) : (
                        <MessageBubble
                          key={item?.id}
                          message={item}
                          onError={setError}
                        />
                      ),
                    )}
                  </div>
                )}
              </div>
            </section>

            <section className="rounded-3xl border border-slate-200 bg-white/90 p-4 shadow-[0_16px_40px_rgba(15,23,42,0.08)] backdrop-blur dark:border-[#ffffff26] dark:bg-[#303030]/80">
              <div className="flex items-end gap-3">
                <input
                  ref={fileInputRef}
                  type="file"
                  multiple
                  onChange={handleFilesSelected}
                  accept=".pdf,.txt,.csv,.xlsx,.xls,application/pdf,text/plain,text/csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
                  className="hidden"
                />
                <button
                  type="button"
                  onClick={() => setSidebarOpen((current) => !current)}
                  aria-expanded={sidebarOpen}
                  aria-controls="chat-sidebar"
                  aria-label={sidebarOpen ? 'Close sidebar' : 'Open sidebar'}
                  className="flex h-11 w-11 items-center justify-center rounded-2xl border border-slate-200 bg-white text-slate-600 shadow-sm transition hover:border-slate-300 hover:text-slate-900 dark:border-[#ffffff26] dark:bg-[#303030] dark:text-[#cdcdcd] dark:hover:border-[#fff3] dark:hover:text-[#ffffff]"
                >
                  <PanelLeft className="h-4 w-4" />
                </button>
                <button
                  type="button"
                  onClick={handleAttachClick}
                  disabled={isLoading}
                  className="flex h-11 w-11 items-center justify-center rounded-2xl border border-slate-200 bg-white text-slate-600 shadow-sm transition hover:border-slate-300 hover:text-slate-900 disabled:cursor-not-allowed disabled:text-slate-300 dark:border-[#ffffff26] dark:bg-[#303030] dark:text-[#cdcdcd] dark:hover:border-[#fff3] dark:hover:text-[#ffffff]"
                  aria-label="Attach files"
                >
                  <Paperclip className="h-4 w-4" />
                </button>
                <textarea
                  ref={inputRef}
                  value={inputValue}
                  onChange={handleInputChange}
                  onKeyDown={handleInputKeyDown}
                  placeholder={`Message ${provider.shortName}...`}
                  rows={1}
                  className="max-h-40 flex-1 resize-none rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm text-slate-700 shadow-sm outline-none transition focus:border-slate-300 focus:ring-2 focus:ring-slate-200 dark:border-[#ffffff26] dark:bg-[#303030] dark:text-[#ffffff] dark:focus:border-[#fff3] dark:focus:ring-[#414141]"
                  disabled={isLoading}
                />
                <button
                  type="button"
                  onClick={handleSend}
                  disabled={isLoading || !canSend}
                  className="flex h-11 w-11 items-center justify-center rounded-2xl bg-[#007AFF] text-white shadow-md transition hover:bg-[#0062d1] disabled:cursor-not-allowed disabled:bg-slate-300"
                  aria-label="Send message"
                >
                  <Send className="h-4 w-4" />
                </button>
              </div>
              {attachments.length ? (
                <div className="mt-3 flex flex-wrap gap-2">
                  {attachments.map((attachment) => (
                    <div
                      key={attachment.id}
                      className="flex items-center gap-2 rounded-full border border-slate-200 bg-white px-3 py-1 text-xs text-slate-600 shadow-sm dark:border-[#ffffff26] dark:bg-[#303030] dark:text-[#f3f3f3]"
                    >
                      <span className="truncate">{attachment.name}</span>
                      <span className="text-[10px] text-slate-400 dark:text-[#afafaf]">
                        {formatFileSize(attachment.size)}
                      </span>
                      <button
                        type="button"
                        onClick={() => handleRemoveAttachment(attachment.id)}
                        className="text-slate-400 transition hover:text-slate-700 dark:text-[#afafaf] dark:hover:text-[#f3f3f3]"
                        aria-label="Remove attachment"
                      >
                        <X className="h-3 w-3" />
                      </button>
                    </div>
                  ))}
                </div>
              ) : null}
              <div className="mt-2 flex flex-wrap items-center justify-between gap-2 text-xs text-slate-500 dark:text-[#afafaf]">
                <span>Enter to send, Shift+Enter for new line</span>
                <span>{inputValue.length} chars</span>
              </div>
            </section>
          </div>
        </main>

        <div
          className={`fixed inset-0 z-40 transition ${
            settingsOpen ? 'pointer-events-auto' : 'pointer-events-none'
          }`}
        >
          <div
            className={`absolute inset-0 bg-slate-900/30 transition-opacity dark:bg-[#181818]/60 ${
              settingsOpen ? 'opacity-100' : 'opacity-0'
            }`}
            onClick={() => setSettingsOpen(false)}
            aria-hidden="true"
          />
          <div
            className={`absolute right-0 top-0 h-full w-full max-w-md transform bg-white shadow-[0_30px_80px_rgba(15,23,42,0.18)] transition-transform dark:bg-[#181818] ${
              settingsOpen ? 'translate-x-0' : 'translate-x-full'
            }`}
            role="dialog"
            aria-modal="true"
          >
            <div className="flex items-center justify-between border-b border-slate-200 px-6 py-4 dark:border-[#ffffff26]">
              <div>
                <div className="text-lg font-semibold text-slate-900 dark:text-[#ffffff]">
                  Settings
                </div>
                <div className="text-xs text-slate-500 dark:text-[#afafaf]">
                  Configure provider access and model settings
                </div>
              </div>
              <button
                type="button"
                onClick={() => setSettingsOpen(false)}
                className="rounded-full border border-slate-200 p-2 text-slate-500 transition hover:text-slate-700 dark:border-[#ffffff26] dark:text-[#afafaf] dark:hover:text-[#f3f3f3]"
                aria-label="Close settings"
              >
                <X className="h-4 w-4" />
              </button>
            </div>

            <div className="flex h-full flex-col gap-6 overflow-y-auto px-6 py-6 pb-28">
              <div>
                <div className="mb-3 text-xs font-semibold uppercase tracking-wide text-slate-500 dark:text-[#afafaf]">
                  Providers
                </div>
                <div className="flex flex-wrap gap-2">
                  {PROVIDER_ORDER.map((providerId) => {
                    const item = PROVIDERS[providerId]
                    const isActive = providerId === activeProvider
                    return (
                      <button
                        key={providerId}
                        type="button"
                        onClick={() => setActiveProvider(providerId)}
                        className={`flex items-center gap-2 rounded-full border px-3 py-1.5 text-xs font-semibold transition ${
                          isActive
                            ? 'border-slate-300 bg-slate-900 text-white dark:border-[#fff3] dark:bg-[#414141] dark:text-[#ffffff]'
                            : 'border-slate-200 bg-white text-slate-600 hover:border-slate-300 dark:border-[#ffffff26] dark:bg-[#303030] dark:text-[#cdcdcd] dark:hover:border-[#fff3]'
                        }`}
                      >
                        <span
                          className="h-2.5 w-2.5 rounded-full"
                          style={{ backgroundColor: item.color }}
                        />
                        <span>{item.name}</span>
                      </button>
                    )
                  })}
                </div>
              </div>

              <div className="space-y-2">
                <div className="text-sm font-semibold text-slate-900 dark:text-[#ffffff]">
                  API key ({provider.name})
                </div>
                <div className="flex items-center gap-2">
                  <input
                    type={showApiKeys[activeProvider] ? 'text' : 'password'}
                    value={apiKeys[activeProvider] ?? ''}
                    onChange={(event) =>
                      setApiKeys((current) => ({
                        ...current,
                        [activeProvider]: event.target.value,
                      }))
                    }
                    placeholder={`Enter ${provider.name} key`}
                    className="flex-1 rounded-2xl border border-slate-200 px-3 py-2 text-sm outline-none transition focus:border-slate-300 focus:ring-2 focus:ring-slate-200 dark:border-[#ffffff26] dark:bg-[#303030] dark:text-[#ffffff] dark:focus:border-[#fff3] dark:focus:ring-[#414141]"
                  />
                  <button
                    type="button"
                    onClick={() =>
                      setShowApiKeys((current) => ({
                        ...current,
                        [activeProvider]: !current[activeProvider],
                      }))
                    }
                    className="rounded-full border border-slate-200 p-2 text-slate-500 transition hover:text-slate-700 dark:border-[#ffffff26] dark:text-[#afafaf] dark:hover:text-[#f3f3f3]"
                    aria-label="Toggle API key visibility"
                  >
                    {showApiKeys[activeProvider] ? (
                      <EyeOff className="h-4 w-4" />
                    ) : (
                      <Eye className="h-4 w-4" />
                    )}
                  </button>
                </div>
                <div className="flex flex-wrap items-center gap-2 text-xs text-slate-500">
                  {apiKeys[activeProvider] ? (
                    <span className="inline-flex items-center gap-1 text-emerald-600">
                      <Check className="h-3 w-3" />
                      Key saved
                    </span>
                  ) : (
                    <span className="inline-flex items-center gap-1 text-amber-600">
                      <AlertTriangle className="h-3 w-3" />
                      Key missing
                    </span>
                  )}
                  {keyWarning ? (
                    <span className="text-amber-600">{keyWarning}</span>
                  ) : null}
                </div>
              </div>

              <div className="space-y-2">
                <div className="text-sm font-semibold text-slate-900 dark:text-[#ffffff]">
                  Model
                </div>
                <select
                  value={activeModel}
                  onChange={(event) => setActiveModel(event.target.value)}
                  className="w-full rounded-2xl border border-slate-200 bg-white px-3 py-2 text-sm outline-none transition focus:border-slate-300 focus:ring-2 focus:ring-slate-200 dark:border-[#ffffff26] dark:bg-[#303030] dark:text-[#ffffff] dark:focus:border-[#fff3] dark:focus:ring-[#414141] [&>option]:bg-white [&>option]:text-slate-900 dark:[&>option]:bg-[#303030] dark:[&>option]:text-[#ffffff]"
                >
                  {provider.models.map((model) => (
                    <option key={model.id} value={model.id}>
                      {model.name}
                    </option>
                  ))}
                </select>
              </div>

              <div className="space-y-2">
                <div className="flex items-center justify-between text-sm font-semibold text-slate-900 dark:text-[#ffffff]">
                  <span>Temperature</span>
                  <span className="text-xs text-slate-500 dark:text-[#afafaf]">
                    {temperature.toFixed(1)}
                  </span>
                </div>
                <input
                  type="range"
                  min={provider.tempRange.min}
                  max={provider.tempRange.max}
                  step={provider.tempRange.step}
                  value={temperature}
                  onChange={(event) =>
                    setTemperature(Number(event.target.value))
                  }
                  disabled={temperatureDisabled}
                  className="w-full accent-slate-900 disabled:cursor-not-allowed disabled:opacity-50 dark:accent-[#dcdcdc]"
                />
                <div className="flex justify-between text-xs text-slate-400 dark:text-[#afafaf]">
                  <span>{provider.tempRange.min}</span>
                  <span>{provider.tempRange.max}</span>
                </div>
                {temperatureDisabled ? (
                  <div className="text-xs text-slate-500 dark:text-[#afafaf]">
                    Temperature is available only with reasoning set to None for
                    GPT-5.2 models.
                  </div>
                ) : null}
              </div>

              {showGpt52Controls ? (
                <div className="space-y-2">
                  <div className="text-sm font-semibold text-slate-900 dark:text-[#ffffff]">
                    Reasoning effort
                  </div>
                  <select
                    value={reasoningEffort}
                    onChange={(event) => setReasoningEffort(event.target.value)}
                    className="w-full rounded-2xl border border-slate-200 bg-white px-3 py-2 text-sm outline-none transition focus:border-slate-300 focus:ring-2 focus:ring-slate-200 dark:border-[#ffffff26] dark:bg-[#303030] dark:text-[#ffffff] dark:focus:border-[#fff3] dark:focus:ring-[#414141] [&>option]:bg-white [&>option]:text-slate-900 dark:[&>option]:bg-[#303030] dark:[&>option]:text-[#ffffff]"
                  >
                    <option value="none">None</option>
                    <option value="low">Low</option>
                    <option value="medium">Medium</option>
                    <option value="high">High</option>
                  </select>
                </div>
              ) : null}

              {showGpt52Controls ? (
                <div className="space-y-2">
                  <div className="text-sm font-semibold text-slate-900 dark:text-[#ffffff]">
                    Verbosity
                  </div>
                  <select
                    value={verbosity}
                    onChange={(event) => setVerbosity(event.target.value)}
                    className="w-full rounded-2xl border border-slate-200 bg-white px-3 py-2 text-sm outline-none transition focus:border-slate-300 focus:ring-2 focus:ring-slate-200 dark:border-[#ffffff26] dark:bg-[#303030] dark:text-[#ffffff] dark:focus:border-[#fff3] dark:focus:ring-[#414141] [&>option]:bg-white [&>option]:text-slate-900 dark:[&>option]:bg-[#303030] dark:[&>option]:text-[#ffffff]"
                  >
                    <option value="low">Low</option>
                    <option value="medium">Medium</option>
                    <option value="high">High</option>
                  </select>
                </div>
              ) : null}

              {showGpt5Controls ? (
                <div className="space-y-2">
                  <div className="text-sm font-semibold text-slate-900 dark:text-[#ffffff]">
                    Reasoning preview
                  </div>
                  <button
                    type="button"
                    onClick={() =>
                      setReasoningPreviewEnabled((value) => !value)
                    }
                    className={`flex w-full items-center justify-between rounded-2xl border px-3 py-2 text-sm transition ${
                      reasoningPreviewEnabled
                        ? 'border-slate-900 bg-slate-900 text-white dark:border-[#fff3] dark:bg-[#414141] dark:text-[#ffffff]'
                        : 'border-slate-200 bg-white text-slate-600 hover:border-slate-300 dark:border-[#ffffff26] dark:bg-[#303030] dark:text-[#cdcdcd] dark:hover:border-[#fff3]'
                    }`}
                  >
                    <span>
                      {reasoningPreviewEnabled ? 'Enabled' : 'Disabled'}
                    </span>
                    <span className="text-xs text-slate-400 dark:text-[#afafaf]">
                      GPT-5 only
                    </span>
                  </button>
                  <div className="text-xs text-slate-500 dark:text-[#afafaf]">
                    Shows a brief, model-provided summary without chain-of-thought.
                  </div>
                </div>
              ) : null}

              <div className="space-y-2">
                <div className="flex items-center justify-between text-sm font-semibold text-slate-900 dark:text-[#ffffff]">
                  <span>Max tokens</span>
                  <span className="text-xs text-slate-500 dark:text-[#afafaf]">
                    {maxTokens.toLocaleString()}
                  </span>
                </div>
                <input
                  type="range"
                  min={provider.maxTokensRange.min}
                  max={provider.maxTokensRange.max}
                  step={1}
                  value={maxTokens}
                  onChange={(event) =>
                    setMaxTokens(
                      clamp(
                        Number(event.target.value),
                        provider.maxTokensRange.min,
                        provider.maxTokensRange.max,
                      ),
                    )
                  }
                  className="w-full accent-slate-900 dark:accent-[#ffffff]"
                />
                <div className="flex items-center justify-between text-[11px] text-slate-400 dark:text-[#afafaf]">
                  <span>{provider.maxTokensRange.min.toLocaleString()}</span>
                  <span>{provider.maxTokensRange.max.toLocaleString()}</span>
                </div>
              </div>

              <div className="space-y-2">
                <div className="flex items-center justify-between text-sm font-semibold text-slate-900 dark:text-[#ffffff]">
                  <span>System prompt</span>
                  <span className="text-xs text-slate-500 dark:text-[#afafaf]">
                    {systemPrompt.length} chars
                  </span>
                </div>
                <textarea
                  value={systemPrompt}
                  onChange={(event) => setSystemPrompt(event.target.value)}
                  rows={4}
                  className="w-full resize-none rounded-2xl border border-slate-200 bg-white px-3 py-2 text-sm outline-none transition focus:border-slate-300 focus:ring-2 focus:ring-slate-200 dark:border-[#ffffff26] dark:bg-[#303030] dark:text-[#ffffff] dark:focus:border-[#fff3] dark:focus:ring-[#414141]"
                />
              </div>

              <div className="flex flex-wrap items-center gap-3">
                <button
                  type="button"
                  onClick={handleClearChat}
                  className="flex items-center gap-2 rounded-full border border-slate-200 px-4 py-2 text-sm font-semibold text-slate-600 transition hover:border-slate-300 hover:text-slate-900 dark:border-[#ffffff26] dark:text-[#cdcdcd] dark:hover:border-[#fff3] dark:hover:text-[#ffffff]"
                >
                  <Trash2 className="h-4 w-4" />
                  Clear chat
                </button>
                <button
                  type="button"
                  onClick={handleExport}
                  disabled={!messages.length}
                  className="flex items-center gap-2 rounded-full border border-slate-200 px-4 py-2 text-sm font-semibold text-slate-600 transition hover:border-slate-300 hover:text-slate-900 disabled:cursor-not-allowed disabled:text-slate-300 dark:border-[#ffffff26] dark:text-[#cdcdcd] dark:hover:border-[#fff3] dark:hover:text-[#ffffff] dark:disabled:text-[#afafaf]"
                >
                  <Download className="h-4 w-4" />
                  Export
                </button>
              </div>
            </div>
          </div>
        </div>

        <ErrorToast message={error} onDismiss={() => setError(null)} />
      </div>
    </div>
  )
}

export default App
