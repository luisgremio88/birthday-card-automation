import { useEffect, useMemo, useState } from 'react'
import './App.css'

const REQUIRED_COLUMNS = ['nome', 'tabelionato', 'email']
const DATE_COLUMNS = ['data_aniversario', 'data_de_nascimento']
const API_BASE = `http://${window.location.hostname}/projeto-aniversario/backend`
const APP_ROOT = `http://${window.location.hostname}/projeto-aniversario`
const DEFAULT_PREVIEW_NAME = 'NOME DO ANIVERSARIANTE'
const DEFAULT_SENDER_EMAIL = 'aniversarios@exemplo.com'
const FIXED_BCC_EMAIL = 'auditoria@exemplo.com'
const SCREENSHOT_MODE = new URLSearchParams(window.location.search).get('shot')
const ASSET_BASE = import.meta.env.BASE_URL

const PROFILES = {
  associado: {
    id: 'associado',
    label: 'Associado',
    excelFile: 'aniversariantes_associado.xlsx',
    templatePath: `${ASSET_BASE}templates/cartao-associado.png`,
    templateHint: 'templates/cartao_base_limpo_associado.png',
    subject: 'Feliz aniversario Associado',
    bodyTitle: 'Feliz aniversario',
    bodyFallback: 'Associado CNB/RS',
    mailDescription:
      'Fluxo para aniversariantes associados, com planilha, PSD anual e texto institucional do associado.',
    nameBox: { x: 179, y: 1538, width: 1043, height: 80, align: 'left' },
  },
  diretoria: {
    id: 'diretoria',
    label: 'Diretoria',
    excelFile: 'aniversariantes_diretoria.xlsx',
    templatePath: `${ASSET_BASE}templates/cartao-diretoria.png`,
    templateHint: 'templates/cartao_base_limpo_diretoria.png',
    subject: 'Parabens ao membro da Diretoria',
    bodyTitle: 'Parabens ao membro da Diretoria',
    bodyFallback: 'Diretoria CNB/RS',
    mailDescription:
      'Fluxo especifico para a Diretoria, com lista propria, layout anual proprio e mensagem dedicada.',
    nameBox: { x: 320, y: 710, width: 1360, height: 115, align: 'center' },
  },
}

const SAMPLE_ROWS = {
  associado: [
    {
      nome: 'Maria Exemplo',
      tabelionato: 'Tabelionato Modelo - Cidade Exemplo',
      email: 'maria.exemplo@empresa.com',
      data_aniversario: '27/04',
    },
    {
      nome: 'Carlos Modelo',
      tabelionato: 'Tabelionato Modelo - Cidade Exemplo',
      email: 'carlos.modelo@empresa.com',
      data_aniversario: '27/04',
    },
  ],
  diretoria: [
    {
      nome: 'Fernanda Exemplo',
      tabelionato: 'Diretoria Institucional',
      email: 'fernanda.exemplo@empresa.com',
      data_aniversario: '27/04',
    },
    {
      nome: 'Roberto Modelo',
      tabelionato: 'Diretoria Institucional',
      email: 'roberto.modelo@empresa.com',
      data_aniversario: '05/05',
    },
  ],
}

const SAMPLE_HISTORY = {
  associado: [
    {
      timestamp: '2026-04-27T08:00:00',
      data_referencia: '27/04/2026',
      nome: 'Maria Exemplo',
      email: 'maria.exemplo@empresa.com',
      status: 'enviado',
      detalhes: 'Cartao enviado com sucesso pelo Outlook.',
      arquivo_cartao_url: null,
    },
    {
      timestamp: '2026-04-27T08:03:00',
      data_referencia: '27/04/2026',
      nome: 'Carlos Modelo',
      email: 'carlos.modelo@empresa.com',
      status: 'rascunho',
      detalhes: 'Rascunho aberto para conferencia.',
      arquivo_cartao_url: null,
    },
  ],
  diretoria: [
    {
      timestamp: '2026-04-27T08:10:00',
      data_referencia: '27/04/2026',
      nome: 'Fernanda Exemplo',
      email: 'fernanda.exemplo@empresa.com',
      status: 'enviado',
      detalhes: 'Parabens ao membro da Diretoria enviado com sucesso.',
      arquivo_cartao_url: null,
    },
  ],
}

const monthNames = [
  'janeiro',
  'fevereiro',
  'marco',
  'abril',
  'maio',
  'junho',
  'julho',
  'agosto',
  'setembro',
  'outubro',
  'novembro',
  'dezembro',
]

const MS_PER_DAY = 24 * 60 * 60 * 1000

function normalizeHeader(value) {
  return String(value ?? '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, '_')
}

function excelDateToString(value, xlsx) {
  if (typeof value === 'number') {
    return xlsx.SSF.format('dd/mm', value)
  }

  const text = String(value ?? '').trim()
  const match = text.match(/^(\d{1,2})[/-](\d{1,2})(?:[/-]\d{2,4})?/)
  if (match) {
    return `${match[1].padStart(2, '0')}/${match[2].padStart(2, '0')}`
  }

  return text
}

function parseExcelDate(value, xlsx) {
  if (value instanceof Date) {
    return value
  }

  if (typeof value === 'number') {
    const parsed = xlsx.SSF.parse_date_code(value)
    if (!parsed) {
      return null
    }
    return new Date(parsed.y, parsed.m - 1, parsed.d)
  }

  const text = String(value ?? '').trim()
  const match = text.match(/^(\d{1,2})[/-](\d{1,2})(?:[/-](\d{2,4}))?$/)
  if (!match) {
    return null
  }

  const day = Number(match[1])
  const month = Number(match[2])
  const year = match[3] ? Number(match[3].length === 2 ? `19${match[3]}` : match[3]) : 2000
  return new Date(year, month - 1, day)
}

function isBlockedBirthDate(value, xlsx) {
  const parsedDate = parseExcelDate(value, xlsx)
  if (!parsedDate) {
    return false
  }

  const day = parsedDate.getDate()
  const month = parsedDate.getMonth() + 1
  const year = parsedDate.getFullYear()

  if (day === 1 && month === 1 && year === 1900) {
    return true
  }

  return parsedDate.getTime() > new Date(2000, 0, 1).getTime()
}

function isDirectorFlag(value) {
  const normalized = String(value ?? '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')

  return ['sim', 's', 'x', '1', 'true', 'diretoria', 'diretor', 'diretores'].includes(normalized)
}

function normalizeNameKey(value) {
  return String(value ?? '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
}

function buildProfileWorkbookFile(xlsx, rows, profileId) {
  const worksheet = xlsx.utils.json_to_sheet(rows)
  const workbook = xlsx.utils.book_new()
  xlsx.utils.book_append_sheet(workbook, worksheet, 'Aniversariantes')
  const arrayBuffer = xlsx.write(workbook, { bookType: 'xlsx', type: 'array' })
  return new File([arrayBuffer], `aniversariantes_${profileId}.xlsx`, {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  })
}

async function parseWorkbookRows(arrayBuffer) {
  const xlsx = await import('xlsx')
  const workbook = xlsx.read(arrayBuffer, { type: 'array' })
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]]
  const rawRows = xlsx.utils.sheet_to_json(firstSheet, {
    defval: '',
    raw: true,
  })

  return rawRows
    .map((rawRow) => {
      const normalizedRow = Object.fromEntries(
        Object.entries(rawRow).map(([key, value]) => [normalizeHeader(key), value]),
      )

      return {
        nome: String(normalizedRow.nome ?? '').trim(),
        tabelionato: String(normalizedRow.tabelionato ?? '').trim(),
        email: String(normalizedRow.email ?? '').trim(),
        data_aniversario: excelDateToString(
          normalizedRow.data_aniversario ?? normalizedRow.data_de_nascimento,
          xlsx,
        ),
      }
    })
    .filter((row) => Object.values(row).some(Boolean))
}

function formatDateLabel(value) {
  const [day, month] = String(value ?? '').split('/')

  if (!day || !month) {
    return 'Data nao informada'
  }

  const monthLabel = monthNames[Number(month) - 1]
  if (!monthLabel) {
    return value
  }

  return `${day.padStart(2, '0')} de ${monthLabel}`
}

function parseBirthdayParts(value) {
  const match = String(value ?? '').trim().match(/^(\d{1,2})[/-](\d{1,2})/)
  if (!match) {
    return null
  }

  const day = Number(match[1])
  const month = Number(match[2])
  if (!Number.isInteger(day) || !Number.isInteger(month) || day < 1 || day > 31 || month < 1 || month > 12) {
    return null
  }

  return { day, month }
}

function startOfToday() {
  const now = new Date()
  return new Date(now.getFullYear(), now.getMonth(), now.getDate())
}

function nextBirthdayInfo(value, today = startOfToday()) {
  const parts = parseBirthdayParts(value)
  if (!parts) {
    return {
      date: null,
      daysUntil: Number.POSITIVE_INFINITY,
      label: 'Data nao informada',
      shortLabel: '',
      isToday: false,
    }
  }

  let nextDate = new Date(today.getFullYear(), parts.month - 1, parts.day)
  if (nextDate < today) {
    nextDate = new Date(today.getFullYear() + 1, parts.month - 1, parts.day)
  }

  const daysUntil = Math.round((nextDate.getTime() - today.getTime()) / MS_PER_DAY)
  const shortLabel = `${String(parts.day).padStart(2, '0')}/${String(parts.month).padStart(2, '0')}`

  return {
    date: nextDate,
    daysUntil,
    label: formatDateLabel(shortLabel),
    shortLabel,
    isToday: daysUntil === 0,
  }
}

function sortRowsByNextBirthday(rowsToSort) {
  const today = startOfToday()

  return rowsToSort
    .map((row, index) => ({
      row,
      originalIndex: index,
      birthday: nextBirthdayInfo(row.data_aniversario, today),
    }))
    .sort((a, b) => {
      if (a.birthday.daysUntil !== b.birthday.daysUntil) {
        return a.birthday.daysUntil - b.birthday.daysUntil
      }

      return String(a.row.nome).localeCompare(String(b.row.nome), 'pt-BR')
    })
}

function formatDaysUntil(daysUntil) {
  if (!Number.isFinite(daysUntil)) {
    return 'Sem data'
  }

  if (daysUntil === 0) {
    return 'Hoje'
  }

  if (daysUntil === 1) {
    return 'Amanha'
  }

  return `Em ${daysUntil} dias`
}

function historyStatusLabel(status) {
  const labels = {
    rascunho: 'Rascunho',
    enviado: 'Enviado',
    erro: 'Erro',
    preparando_envio: 'Preparando',
    ignorado: 'Ignorado',
  }

  return labels[status] || status || 'Sem status'
}

function fitSingleLineName(context, text, maxWidth) {
  for (let fontSize = 88; fontSize >= 24; fontSize -= 1) {
    context.font = `700 ${fontSize}px Tahoma`
    if (context.measureText(text).width <= maxWidth) {
      return fontSize
    }
  }

  return 10
}

function loadImage(source) {
  return new Promise((resolve, reject) => {
    const image = new Image()
    image.onload = () => resolve(image)
    image.onerror = reject
    image.src = source
  })
}

async function renderBirthdayCard(name, profile, templateVersion) {
  const image = await loadImage(`${profile.templatePath}?v=${templateVersion}`)
  const canvas = document.createElement('canvas')
  const context = canvas.getContext('2d')
  const safeName = String(name || DEFAULT_PREVIEW_NAME).trim().toUpperCase()
  const nameBox = profile.nameBox

  canvas.width = image.width
  canvas.height = image.height

  context.drawImage(image, 0, 0)
  const fontSize = fitSingleLineName(context, safeName, nameBox.width)
  context.font = `700 ${fontSize}px Tahoma`
  const metrics = context.measureText(safeName)
  const textHeight =
    (metrics.actualBoundingBoxAscent || fontSize * 0.75) +
    (metrics.actualBoundingBoxDescent || fontSize * 0.2)
  const baselineY =
    nameBox.y + (nameBox.height - textHeight) / 2 + (metrics.actualBoundingBoxAscent || fontSize * 0.75)

  context.fillStyle = '#ffffff'
  context.textAlign = nameBox.align === 'center' ? 'center' : 'left'
  context.textBaseline = 'alphabetic'
  const textX = nameBox.align === 'center' ? nameBox.x + nameBox.width / 2 : nameBox.x
  context.fillText(safeName, textX, baselineY)

  return canvas.toDataURL('image/png')
}

async function parseJsonResponse(response) {
  const text = await response.text()
  if (!text.trim()) {
    throw new Error('O servidor respondeu vazio. Tente novamente e verifique se o Apache e o Outlook estao abertos.')
  }

  try {
    return JSON.parse(text)
  } catch {
    throw new Error(text.slice(0, 240) || 'O servidor retornou uma resposta invalida.')
  }
}

function App() {
  const isScreenshotMode = Boolean(SCREENSHOT_MODE)
  const [currentProfileId, setCurrentProfileId] = useState('associado')
  const [rows, setRows] = useState(isScreenshotMode ? SAMPLE_ROWS.associado : [])
  const [fileName, setFileName] = useState(
    isScreenshotMode ? 'aniversariantes_associado.xlsx' : '',
  )
  const [error, setError] = useState('')
  const [selectedRowIndex, setSelectedRowIndex] = useState(isScreenshotMode ? 0 : null)
  const [previewName, setPreviewName] = useState(
    isScreenshotMode ? SAMPLE_ROWS.associado[0].nome : DEFAULT_PREVIEW_NAME,
  )
  const [previewCardUrl, setPreviewCardUrl] = useState('')
  const [templateVersion, setTemplateVersion] = useState(0)
  const [senderEmail, setSenderEmail] = useState(DEFAULT_SENDER_EMAIL)
  const [fixedBccEmail, setFixedBccEmail] = useState(FIXED_BCC_EMAIL)
  const [uploadStatus, setUploadStatus] = useState(
    isScreenshotMode ? 'Planilha salva com sucesso no perfil associado.' : '',
  )
  const [templateStatus, setTemplateStatus] = useState(
    isScreenshotMode ? 'Template PSD processado e base limpa atualizada.' : '',
  )
  const [automationStatus, setAutomationStatus] = useState(
    isScreenshotMode ? 'Fluxo demonstrativo pronto para abrir rascunho ou enviar.' : '',
  )
  const [automationOutput, setAutomationOutput] = useState(
    isScreenshotMode ? ['Maria Exemplo <maria.exemplo@empresa.com>: enviado'] : [],
  )
  const [isRunningAutomation, setIsRunningAutomation] = useState(false)
  const [historyItems, setHistoryItems] = useState(
    isScreenshotMode ? SAMPLE_HISTORY.associado : [],
  )
  const [historyStatus, setHistoryStatus] = useState(
    isScreenshotMode ? '' : 'Carregando historico...',
  )
  const [isHistoryOpen, setIsHistoryOpen] = useState(isScreenshotMode)
  const [isBirthdayListOpen, setIsBirthdayListOpen] = useState(isScreenshotMode)
  const [editingRowIndex, setEditingRowIndex] = useState(null)
  const [editingRow, setEditingRow] = useState(null)

  const currentProfile = PROFILES[currentProfileId]

  const stats = useMemo(() => {
    const total = rows.length
    const uniqueUnits = new Set(rows.map((row) => row.tabelionato).filter(Boolean)).size
    const today = startOfToday()
    const todayKey = `${String(today.getDate()).padStart(2, '0')}/${String(today.getMonth() + 1).padStart(2, '0')}`
    const birthdaysToday = rows.filter(
      (row) => String(row.data_aniversario).slice(0, 5) === todayKey,
    ).length
    const upcomingInSevenDays = rows.filter((row) => {
      const birthday = nextBirthdayInfo(row.data_aniversario, today)
      return birthday.daysUntil >= 0 && birthday.daysUntil <= 7
    }).length

    return { total, uniqueUnits, birthdaysToday, todayKey, upcomingInSevenDays }
  }, [rows])

  const sortedBirthdayRows = useMemo(() => sortRowsByNextBirthday(rows), [rows])
  const nextBirthday = sortedBirthdayRows[0] || null
  const upcomingRows = sortedBirthdayRows.slice(0, 5)

  const selectedRow = selectedRowIndex !== null ? rows[selectedRowIndex] ?? null : null
  const recipientEmail = selectedRow?.email || ''
  const bccRecipients = [recipientEmail, fixedBccEmail].filter(Boolean).join('; ')
  const emailSubject = currentProfile.subject
  const emailBodyTitle = selectedRow
    ? `${currentProfile.bodyTitle}, ${selectedRow.nome}!`
    : `${currentProfile.bodyTitle}!`
  const emailBodyUnit = selectedRow?.tabelionato || currentProfile.bodyFallback

  useEffect(() => {
    if (isScreenshotMode) {
      return undefined
    }

    let active = true

    async function loadConfig() {
      try {
        const response = await fetch(`${API_BASE}/config.php`)
        const payload = await parseJsonResponse(response)
        if (!response.ok || !payload.success || !active) {
          return
        }

        setSenderEmail(payload.emailDefaults?.senderEmail || DEFAULT_SENDER_EMAIL)
        setFixedBccEmail(payload.emailDefaults?.bccEmail || FIXED_BCC_EMAIL)
      } catch {
        if (!active) {
          return
        }
      }
    }

    loadConfig()

    return () => {
      active = false
    }
  }, [isScreenshotMode])

  useEffect(() => {
    let active = true

    renderBirthdayCard(previewName, currentProfile, templateVersion)
      .then((imageUrl) => {
        if (active) {
          setPreviewCardUrl(imageUrl)
        }
      })
      .catch(() => {
        if (active) {
          setPreviewCardUrl('')
        }
      })

    return () => {
      active = false
    }
  }, [previewName, currentProfile, templateVersion])

  useEffect(() => {
    if (isScreenshotMode) {
      return undefined
    }

    let active = true

    async function loadHistory() {
      setHistoryStatus('Carregando historico...')

      try {
        const response = await fetch(`${API_BASE}/history.php?profile=${currentProfileId}`)
        const payload = await response.json()

        if (!response.ok || !payload.success) {
          throw new Error(payload.message || 'Falha ao carregar o historico.')
        }

        if (!active) {
          return
        }

        setHistoryItems(payload.items || [])
        setHistoryStatus(payload.items?.length ? '' : 'Nenhum envio registrado ainda para este perfil.')
      } catch (historyError) {
        if (!active) {
          return
        }

        setHistoryItems([])
        setHistoryStatus(historyError.message || 'Falha ao carregar o historico.')
      }
    }

    loadHistory()

    return () => {
      active = false
    }
  }, [currentProfileId, isScreenshotMode])

  useEffect(() => {
    if (isScreenshotMode) {
      return undefined
    }

    let active = true

    async function loadSavedSpreadsheet() {
      const spreadsheetUrl = `${APP_ROOT}/uploads/${currentProfile.excelFile}?v=${Date.now()}`

      try {
        const response = await fetch(spreadsheetUrl)
        if (!response.ok) {
          return
        }

        const rowsFromServer = await parseWorkbookRows(await response.arrayBuffer())
        if (!active) {
          return
        }

        setRows(rowsFromServer)
        setFileName(currentProfile.excelFile)
        setSelectedRowIndex(rowsFromServer.length > 0 ? 0 : null)
        setPreviewName(rowsFromServer[0]?.nome || DEFAULT_PREVIEW_NAME)
        setUploadStatus(
          rowsFromServer.length > 0
            ? `Planilha salva carregada: ${rowsFromServer.length} aniversariante(s).`
            : '',
        )
      } catch {
        if (!active) {
          return
        }
      }
    }

    loadSavedSpreadsheet()

    return () => {
      active = false
    }
  }, [currentProfile, currentProfileId, isScreenshotMode])

  function resetStatuses() {
    setError('')
    setUploadStatus('')
    setTemplateStatus('')
    setAutomationStatus('')
    setAutomationOutput([])
  }

  function handleProfileChange(profileId) {
    setCurrentProfileId(profileId)
    if (isScreenshotMode) {
      setRows(SAMPLE_ROWS[profileId])
      setFileName(`aniversariantes_${profileId}.xlsx`)
      setSelectedRowIndex(0)
      setPreviewName(SAMPLE_ROWS[profileId][0]?.nome || DEFAULT_PREVIEW_NAME)
      setHistoryItems(SAMPLE_HISTORY[profileId] || [])
      setHistoryStatus('')
    } else {
      setRows([])
      setFileName('')
      setSelectedRowIndex(null)
      setPreviewName(DEFAULT_PREVIEW_NAME)
      setHistoryItems([])
      setHistoryStatus('Carregando historico...')
    }
    setPreviewCardUrl('')
    setIsHistoryOpen(isScreenshotMode)
    setIsBirthdayListOpen(isScreenshotMode)
    setEditingRowIndex(null)
    setEditingRow(null)
    resetStatuses()
  }

  async function saveExcelOnServer(file, profileId = currentProfileId) {
    const formData = new FormData()
    formData.append('excel', file)
    formData.append('profile', profileId)

    const response = await fetch(`${API_BASE}/upload_excel.php`, {
      method: 'POST',
      body: formData,
    })
    const payload = await response.json()

    if (!response.ok || !payload.success) {
      throw new Error(payload.message || 'Nao foi possivel salvar a planilha no servidor.')
    }

    return payload
  }

  async function persistRows(nextRows) {
    const xlsx = await import('xlsx')
    const file = buildProfileWorkbookFile(xlsx, nextRows, currentProfileId)
    return saveExcelOnServer(file, currentProfileId)
  }

  async function saveTemplateOnServer(file) {
    const formData = new FormData()
    formData.append('templatePsd', file)
    formData.append('profile', currentProfileId)

    const response = await fetch(`${API_BASE}/upload_template.php`, {
      method: 'POST',
      body: formData,
    })
    const payload = await response.json()

    if (!response.ok || !payload.success) {
      throw new Error(payload.message || 'Nao foi possivel atualizar o template PSD.')
    }

    return payload
  }

  function handleFileChange(event) {
    const file = event.target.files?.[0]
    if (!file) {
      return
    }

    setError('')
    setFileName(file.name)
    setUploadStatus('Lendo planilha, removendo duplicidades e separando Associado/Diretoria...')
    setTemplateStatus('')
    setAutomationStatus('')

    const reader = new FileReader()
    reader.onload = async (loadEvent) => {
      try {
        const xlsx = await import('xlsx')
        const workbook = xlsx.read(loadEvent.target?.result, { type: 'array' })
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]]
        const rawRows = xlsx.utils.sheet_to_json(firstSheet, {
          defval: '',
          raw: true,
        })

        if (rawRows.length === 0) {
          setRows([])
          setSelectedRowIndex(null)
          setPreviewName(DEFAULT_PREVIEW_NAME)
          setError('A planilha esta vazia. Adicione os aniversariantes e tente novamente.')
          return
        }

        const originalHeaders = Object.keys(rawRows[0] ?? {}).map(normalizeHeader)
        const absentHeaders = REQUIRED_COLUMNS.filter((column) => !originalHeaders.includes(column))
        const hasBirthDateColumn = DATE_COLUMNS.some((column) => originalHeaders.includes(column))

        if (absentHeaders.length > 0 || !hasBirthDateColumn) {
          setRows([])
          setSelectedRowIndex(null)
          setPreviewName(DEFAULT_PREVIEW_NAME)
          const missingColumns = [...absentHeaders]
          if (!hasBirthDateColumn) {
            missingColumns.push('data_aniversario ou data_de_nascimento')
          }
          setError(`Colunas obrigatorias ausentes: ${missingColumns.join(', ')}`)
          return
        }

        const profileRows = { associado: [], diretoria: [] }
        const seenNames = new Set()
        let duplicateCount = 0
        let invalidDateCount = 0

        rawRows.forEach((rawRow) => {
          const normalizedRow = Object.fromEntries(
            Object.entries(rawRow).map(([key, value]) => [normalizeHeader(key), value]),
          )
          const name = String(normalizedRow.nome ?? '').trim()
          const nameKey = normalizeNameKey(name)

          if (!nameKey) {
            return
          }

          if (seenNames.has(nameKey)) {
            duplicateCount += 1
            return
          }
          seenNames.add(nameKey)

          const birthDate = normalizedRow.data_de_nascimento ?? normalizedRow.data_aniversario
          if (isBlockedBirthDate(birthDate, xlsx)) {
            invalidDateCount += 1
            return
          }

          const row = {
            nome: name,
            tabelionato: String(normalizedRow.tabelionato ?? '').trim(),
            email: String(normalizedRow.email ?? '').trim(),
            data_aniversario: excelDateToString(birthDate, xlsx),
          }

          if (!row.email || !row.data_aniversario) {
            return
          }

          const profileId = isDirectorFlag(normalizedRow.diretores) ? 'diretoria' : 'associado'
          profileRows[profileId].push(row)
        })

        const visibleRows = profileRows[currentProfileId]
        setRows(visibleRows)
        setSelectedRowIndex(visibleRows.length > 0 ? 0 : null)
        setPreviewName(visibleRows[0]?.nome || DEFAULT_PREVIEW_NAME)

        await Promise.all(
          Object.entries(profileRows).map(([profileId, profileList]) =>
            saveExcelOnServer(buildProfileWorkbookFile(xlsx, profileList, profileId), profileId),
          ),
        )

        setUploadStatus(
          `Planilha processada: ${profileRows.associado.length} associado(s), ${profileRows.diretoria.length} diretoria. ` +
            `${duplicateCount} duplicado(s) ignorado(s), ${invalidDateCount} data(s) invalida(s) ignorada(s).`,
        )
      } catch (fileError) {
        setRows([])
        setSelectedRowIndex(null)
        setPreviewName(DEFAULT_PREVIEW_NAME)
        setUploadStatus('')
        setError(
          fileError.message ||
            'Nao foi possivel ler ou salvar o arquivo. Verifique se ele esta no formato .xlsx ou .xls.',
        )
      }
    }

    reader.readAsArrayBuffer(file)
  }

  async function handleTemplateChange(event) {
    const file = event.target.files?.[0]
    if (!file) {
      return
    }

    setError('')
    setTemplateStatus(`Enviando PSD do perfil ${currentProfile.label.toLowerCase()} e regenerando a base limpa...`)

    try {
      const result = await saveTemplateOnServer(file)
      setTemplateStatus(result.message || 'Template PSD atualizado com sucesso.')
      setTemplateVersion(Date.now())
    } catch (templateError) {
      setTemplateStatus('')
      setError(templateError.message || 'Nao foi possivel atualizar o template PSD.')
    }
  }

  function handleSelectRow(index) {
    const row = rows[index]
    setSelectedRowIndex(index)
    setPreviewName(row?.nome || DEFAULT_PREVIEW_NAME)
  }

  function handleStartEdit(index) {
    setEditingRowIndex(index)
    setEditingRow({ ...rows[index] })
    setError('')
  }

  function handleCancelEdit() {
    setEditingRowIndex(null)
    setEditingRow(null)
  }

  function handleEditField(field, value) {
    setEditingRow((current) => ({
      ...(current || {}),
      [field]: value,
    }))
  }

  async function handleSaveEdit(index) {
    if (!editingRow?.nome?.trim() || !editingRow?.email?.trim() || !editingRow?.data_aniversario?.trim()) {
      setError('Nome, e-mail e data de aniversario sao obrigatorios para salvar.')
      return
    }

    const nextRows = rows.map((row, rowIndex) => (rowIndex === index ? editingRow : row))
    setRows(nextRows)
    setSelectedRowIndex(index)
    setPreviewName(editingRow.nome)
    setEditingRowIndex(null)
    setEditingRow(null)

    try {
      await persistRows(nextRows)
      setUploadStatus(`Alteracao salva na planilha ${currentProfile.label}.`)
      setError('')
    } catch (saveError) {
      setError(saveError.message || 'Nao foi possivel salvar a alteracao na planilha.')
    }
  }

  async function handleDeleteRow(index) {
    const row = rows[index]
    const confirmed = window.confirm(`Excluir ${row?.nome || 'este aniversariante'} da lista ${currentProfile.label}?`)
    if (!confirmed) {
      return
    }

    const nextRows = rows.filter((_, rowIndex) => rowIndex !== index)
    setRows(nextRows)
    setSelectedRowIndex(nextRows.length > 0 ? Math.min(index, nextRows.length - 1) : null)
    setPreviewName(nextRows[Math.min(index, nextRows.length - 1)]?.nome || DEFAULT_PREVIEW_NAME)
    setEditingRowIndex(null)
    setEditingRow(null)

    try {
      await persistRows(nextRows)
      setUploadStatus(`Registro excluido e planilha ${currentProfile.label} salva.`)
      setError('')
    } catch (deleteError) {
      setError(deleteError.message || 'Nao foi possivel salvar a exclusao na planilha.')
    }
  }

  async function runAutomation(mode) {
    if (isScreenshotMode) {
      setAutomationStatus(
        mode === 'send'
          ? 'Demo: envio automatico exibido apenas para captura do portfolio.'
          : 'Demo: rascunho do Outlook simulado para a captura do portfolio.',
      )
      setAutomationOutput(
        mode === 'send'
          ? ['Maria Exemplo <maria.exemplo@empresa.com>: enviado']
          : ['Maria Exemplo <maria.exemplo@empresa.com>: rascunho'],
      )
      return
    }

    if (mode === 'send') {
      const confirmed = window.confirm(
        `Enviar agora os e-mails de aniversario do perfil ${currentProfile.label}? O sistema vai ignorar quem ja constar como enviado no historico do dia.`,
      )
      if (!confirmed) {
        return
      }
    }

    setIsRunningAutomation(true)
    setAutomationStatus(
      mode === 'send'
        ? `Enviando e-mails do perfil ${currentProfile.label.toLowerCase()}...`
        : `Abrindo rascunhos do perfil ${currentProfile.label.toLowerCase()} no Outlook...`,
    )
    setAutomationOutput([])

    try {
      const response = await fetch(`${API_BASE}/run_automation.php`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          profile: currentProfileId,
          mode,
          senderEmail,
          bccEmail: fixedBccEmail,
        }),
      })

      const payload = await parseJsonResponse(response)

      if (!response.ok || !payload.success) {
        setAutomationOutput(payload.output || [])
        throw new Error(payload.message || 'Falha ao executar a automacao.')
      }

      setAutomationStatus(payload.message || 'Automacao executada com sucesso.')
      setAutomationOutput(payload.output || [])
      const historyResponse = await fetch(`${API_BASE}/history.php?profile=${currentProfileId}`)
      const historyPayload = await historyResponse.json()
      if (historyResponse.ok && historyPayload.success) {
        setHistoryItems(historyPayload.items || [])
        setHistoryStatus(historyPayload.items?.length ? '' : 'Nenhum envio registrado ainda para este perfil.')
      }
    } catch (automationError) {
      setAutomationStatus(automationError.message || 'Falha ao executar a automacao.')
    } finally {
      setIsRunningAutomation(false)
    }
  }

  const showOverviewSection = !SCREENSHOT_MODE || SCREENSHOT_MODE === 'overview'
  const showPreviewSection = !SCREENSHOT_MODE || SCREENSHOT_MODE === 'preview'
  const showHistorySection = !SCREENSHOT_MODE || SCREENSHOT_MODE === 'history'

  return (
    <main className="page-shell">
      {showOverviewSection ? (
        <>
      <section className="hero-panel">
        <div className="hero-copy">
          <span className="eyebrow">Painel de envio</span>
          <h1>Sistema de Cartoes de Aniversario</h1>
          <p className="lead">
            Importe a planilha, confira os aniversariantes do dia e prepare o envio pelo Outlook.
          </p>

          <div className="profile-switcher">
            {Object.values(PROFILES).map((profile) => (
              <button
                key={profile.id}
                type="button"
                className={currentProfileId === profile.id ? 'profile-pill active' : 'profile-pill'}
                onClick={() => handleProfileChange(profile.id)}
              >
                {profile.label}
              </button>
            ))}
          </div>

          <div className="hero-actions">
            <label className="upload-button" htmlFor="excel-upload">
              Importar Excel {currentProfile.label}
            </label>
            <input id="excel-upload" type="file" accept=".xlsx,.xls" onChange={handleFileChange} />

            <button
              type="button"
              className="ghost-button"
              onClick={() => setPreviewName(selectedRow?.nome || DEFAULT_PREVIEW_NAME)}
            >
              Gerar cartao teste
            </button>

            <label className="secondary-upload" htmlFor="template-upload">
              Anexar PSD {currentProfile.label}
            </label>
            <input id="template-upload" type="file" accept=".psd" onChange={handleTemplateChange} />
          </div>

          <div className="template-card">
            <p>Perfil atual: {currentProfile.label}</p>
            <code>nome | tabelionato | email | data_aniversario</code>
            <span>Formato sugerido para a data: 16/04</span>
            <span>Template carregado: `{currentProfile.templateHint}`</span>
            <span>{currentProfile.mailDescription}</span>
          </div>
        </div>

        <aside className="status-card">
          <p className="status-label">Arquivo atual</p>
          <strong>{fileName || `Nenhuma planilha ${currentProfile.label.toLowerCase()} importada`}</strong>
          <p className="status-helper">
            O sistema le a primeira aba do Excel e organiza os aniversariantes do perfil{' '}
            {currentProfile.label.toLowerCase()} para as proximas etapas.
          </p>
          {uploadStatus ? <p className="success-banner">{uploadStatus}</p> : null}
          {templateStatus ? <p className="success-banner">{templateStatus}</p> : null}
          {error ? <p className="error-banner">{error}</p> : null}
        </aside>
      </section>

      <section className="stats-grid">
        <article className="stat-card">
          <span>Total de registros</span>
          <strong>{stats.total}</strong>
        </article>
        <article className="stat-card">
          <span>Tabelionatos</span>
          <strong>{stats.uniqueUnits}</strong>
        </article>
        <article className="stat-card">
          <span>Aniversariantes de hoje</span>
          <strong>{stats.birthdaysToday}</strong>
          <small>Comparacao com {stats.todayKey}</small>
        </article>
        <article className="stat-card next-stat-card">
          <span>Proximos 7 dias</span>
          <strong>{stats.upcomingInSevenDays}</strong>
          <small>{nextBirthday ? `Proximo: ${nextBirthday.row.nome}` : 'Nenhuma planilha carregada'}</small>
        </article>
      </section>

      {nextBirthday ? (
        <section className="upcoming-panel top-upcoming-panel">
          <article className="next-birthday-card">
            <span>Proximo envio</span>
            <strong>{nextBirthday.row.nome}</strong>
            <p>{nextBirthday.row.tabelionato || currentProfile.bodyFallback}</p>
            <div>
              <b>{nextBirthday.birthday.label}</b>
              <small>{formatDaysUntil(nextBirthday.birthday.daysUntil)}</small>
            </div>
          </article>

          <div className="upcoming-list">
            {upcomingRows.map((item) => (
              <button
                key={`${currentProfileId}-upcoming-${item.originalIndex}-${item.row.email}`}
                type="button"
                className={item.birthday.isToday ? 'upcoming-item today' : 'upcoming-item'}
                onClick={() => handleSelectRow(item.originalIndex)}
              >
                <span>{item.row.nome}</span>
                <strong>{item.birthday.shortLabel}</strong>
                <small>{formatDaysUntil(item.birthday.daysUntil)}</small>
              </button>
            ))}
          </div>
        </section>
      ) : null}
        </>
      ) : null}

      {showPreviewSection ? (
        <>
      <section className="preview-panel">
        <div className="preview-copy">
          <div>
            <span className="eyebrow">Geracao do cartao</span>
            <h2>Previa automatica {currentProfile.label}</h2>
          </div>
          <p>
            O sistema usa a base limpa extraida do PSD do perfil {currentProfile.label.toLowerCase()}
            {' '}e escreve o nome automaticamente na faixa correta da arte.
          </p>

          <label className="field-label" htmlFor="preview-name">
            Nome no cartao
          </label>
          <input
            id="preview-name"
            className="text-input"
            type="text"
            value={previewName}
            onChange={(event) => setPreviewName(event.target.value)}
            placeholder="Digite o nome do aniversariante"
          />

          {selectedRow ? (
            <div className="selected-card">
              <strong>{selectedRow.nome}</strong>
              <span>{selectedRow.tabelionato}</span>
              <span>{selectedRow.email}</span>
            </div>
          ) : (
            <div className="selected-card">
              <strong>Cartao de exemplo</strong>
              <span>Importe a planilha do perfil atual para selecionar um aniversariante real.</span>
            </div>
          )}
        </div>

        <div className="preview-frame">
          {previewCardUrl ? (
            <img src={previewCardUrl} alt="Previa do cartao de aniversario gerado automaticamente" />
          ) : (
            <div className="preview-placeholder">Nao foi possivel gerar a previa do cartao.</div>
          )}
        </div>
      </section>

      <section className="mail-panel">
        <div className="mail-composer">
          <div className="mail-row">
            <div className="mail-tag">De</div>
            <input
              className="mail-input"
              type="email"
              value={senderEmail}
              onChange={(event) => setSenderEmail(event.target.value)}
              placeholder="Digite o e-mail remetente"
            />
          </div>

          <div className="mail-row">
            <div className="mail-tag active">Para</div>
            <input
              className="mail-input"
              type="email"
              value={senderEmail}
              readOnly
              placeholder="O Outlook usa um destinatario tecnico no campo Para"
            />
          </div>

          <div className="mail-row">
            <div className="mail-tag">Cc</div>
            <input className="mail-input" type="text" value="" readOnly placeholder="Opcional para uma proxima etapa" />
          </div>

          <div className="mail-row">
            <div className="mail-tag">Cco</div>
            <input
              className="mail-input"
              type="email"
              value={fixedBccEmail}
              onChange={(event) => setFixedBccEmail(event.target.value)}
              placeholder="Digite o e-mail fixo oculto"
            />
          </div>

          <div className="mail-row subject-row">
            <div className="mail-label">Assunto</div>
            <input className="mail-input" type="text" value={emailSubject} readOnly />
          </div>

          <div className="mail-body">
            <p className="mail-body-title">{emailBodyTitle}</p>
            <p className="mail-body-text">{emailBodyUnit}</p>
            {previewCardUrl ? (
              <img
                src={previewCardUrl}
                alt="Cartao que sera enviado no corpo do e-mail"
                className="mail-card-image"
              />
            ) : (
              <div className="preview-placeholder">O cartao sera exibido aqui dentro do corpo do e-mail.</div>
            )}
          </div>
        </div>

        <aside className="mail-sidecard">
          <span className="eyebrow">Fluxo do envio</span>
          <h2>Preparacao do e-mail {currentProfile.label}</h2>
          <p>
            O campo <strong>De</strong> usa a conta configurada no Outlook. O campo
            <strong> Para</strong> fica tecnico, enquanto os destinatarios reais seguem ocultos no
            <strong> Cco</strong>.
          </p>
          <p>
            O sistema envia em <code>Cco</code> o e-mail da planilha e tambem o e-mail fixo
            configurado por voce. Destinatarios previstos: <code>{bccRecipients || 'nenhum'}</code>.
          </p>

          <div className="automation-actions">
            <button
              type="button"
              className="primary-action"
              onClick={() => runAutomation('draft')}
              disabled={isRunningAutomation}
            >
              Abrir rascunho no Outlook
            </button>
            <button
              type="button"
              className="secondary-action"
              onClick={() => runAutomation('send')}
              disabled={isRunningAutomation}
            >
              Enviar e-mails do dia
            </button>
          </div>

          {automationStatus ? <p className="status-note">{automationStatus}</p> : null}
          {automationOutput.length > 0 ? (
            <div className="automation-log">
              {automationOutput.map((line) => (
                <p key={line}>{line}</p>
              ))}
            </div>
          ) : null}
        </aside>
      </section>
        </>
      ) : null}

      {showHistorySection ? (
        <>
      <section className="history-toggle-panel">
        <div>
          <span className="eyebrow">Historico</span>
          <h2>Envios do perfil {currentProfile.label}</h2>
          <p>
            Consulte somente quando quiser revisar rascunhos, envios concluídos ou erros do Outlook.
          </p>
        </div>
        <button
          type="button"
          className="history-toggle-button"
          onClick={() => setIsHistoryOpen((current) => !current)}
        >
          {isHistoryOpen ? 'Ocultar historico' : `Ver historico (${historyItems.length})`}
        </button>
      </section>

      {isHistoryOpen ? (
        <section className="history-panel">
        <div className="table-header">
          <div>
            <span className="eyebrow">Historico</span>
            <h2>Envios do perfil {currentProfile.label}</h2>
          </div>
          <p>
            Aqui voce acompanha quando o sistema abriu rascunhos, enviou e-mails ou encontrou
            algum erro no fluxo do perfil atual.
          </p>
        </div>

        {historyItems.length > 0 ? (
          <div className="table-wrapper">
            <table>
              <thead>
                <tr>
                  <th>Data</th>
                  <th>Nome</th>
                  <th>E-mail</th>
                  <th>Status</th>
                  <th>Cartao</th>
                  <th>Detalhes</th>
                </tr>
              </thead>
              <tbody>
                {historyItems.map((item) => (
                  <tr key={`${item.timestamp}-${item.email}`}>
                    <td>{item.data_referencia}</td>
                    <td>{item.nome}</td>
                    <td>{item.email}</td>
                    <td>
                      <span className={`status-chip status-${item.status}`}>
                        {historyStatusLabel(item.status)}
                      </span>
                    </td>
                    <td>
                      {item.arquivo_cartao_url ? (
                        <a
                          className="history-link"
                          href={`http://${window.location.hostname}${item.arquivo_cartao_url}`}
                          target="_blank"
                          rel="noreferrer"
                        >
                          Abrir cartao
                        </a>
                      ) : (
                        'Sem arquivo'
                      )}
                    </td>
                    <td>{item.detalhes}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ) : (
          <div className="empty-state">
            <h3>Historico do perfil {currentProfile.label}</h3>
            <p>{historyStatus}</p>
          </div>
        )}
      </section>
      ) : null}

      <section className="list-toggle-panel">
          <div>
            <span className="eyebrow">Visualizacao</span>
            <h2>Lista de aniversariantes {currentProfile.label}</h2>
          <p>
            {rows.length > 0
              ? `${rows.length} registro(s) carregado(s). Abra a lista completa para editar, excluir ou usar no cartao.`
              : `Importe o Excel do perfil ${currentProfile.label.toLowerCase()} para visualizar os aniversariantes.`}
          </p>
        </div>
        <button
          type="button"
          className="history-toggle-button"
          onClick={() => setIsBirthdayListOpen((current) => !current)}
        >
          {isBirthdayListOpen ? 'Ocultar lista' : `Ver lista (${rows.length})`}
        </button>
      </section>

      {isBirthdayListOpen ? (
        <section className="table-panel">
          <div className="table-header">
            <div>
              <span className="eyebrow">Visualizacao</span>
              <h2>Lista de aniversariantes {currentProfile.label}</h2>
            </div>
            <p>
              {rows.length > 0
                ? `Dados lidos com sucesso para o perfil ${currentProfile.label.toLowerCase()}.`
                : `Importe o Excel do perfil ${currentProfile.label.toLowerCase()} para visualizar os aniversariantes aqui.`}
            </p>
          </div>

        {rows.length > 0 ? (
          <div className="table-wrapper">
            <table>
              <thead>
                <tr>
                  <th>Nome</th>
                  <th>Tabelionato</th>
                  <th>E-mail</th>
                  <th>Data de aniversario</th>
                  <th>Proximo envio</th>
                  <th>Cartao</th>
                  <th>Acoes</th>
                </tr>
              </thead>
              <tbody>
                {sortedBirthdayRows.map((item) => {
                  const row = item.row
                  const index = item.originalIndex
                  const isEditing = editingRowIndex === index
                  const displayRow = isEditing ? editingRow : row

                  return (
                    <tr
                      key={`${currentProfileId}-${row.email}-${index}`}
                      className={[
                        selectedRowIndex === index ? 'selected-row' : '',
                        item.birthday.isToday ? 'birthday-today-row' : '',
                      ]
                        .filter(Boolean)
                        .join(' ')}
                    >
                      <td>
                        {isEditing ? (
                          <input
                            className="table-input"
                            value={displayRow?.nome || ''}
                            onChange={(event) => handleEditField('nome', event.target.value)}
                          />
                        ) : (
                          row.nome || 'Nao informado'
                        )}
                      </td>
                      <td>
                        {isEditing ? (
                          <input
                            className="table-input"
                            value={displayRow?.tabelionato || ''}
                            onChange={(event) => handleEditField('tabelionato', event.target.value)}
                          />
                        ) : (
                          row.tabelionato || 'Nao informado'
                        )}
                      </td>
                      <td>
                        {isEditing ? (
                          <input
                            className="table-input"
                            type="email"
                            value={displayRow?.email || ''}
                            onChange={(event) => handleEditField('email', event.target.value)}
                          />
                        ) : (
                          row.email || 'Nao informado'
                        )}
                      </td>
                      <td>
                        {isEditing ? (
                          <input
                            className="table-input table-date-input"
                            value={displayRow?.data_aniversario || ''}
                            onChange={(event) => handleEditField('data_aniversario', event.target.value)}
                            placeholder="DD/MM"
                          />
                        ) : (
                          formatDateLabel(row.data_aniversario)
                        )}
                      </td>
                      <td>
                        <span className={item.birthday.isToday ? 'next-date-chip today' : 'next-date-chip'}>
                          {formatDaysUntil(item.birthday.daysUntil)}
                        </span>
                      </td>
                      <td>
                        <button
                          type="button"
                          className="table-button"
                          onClick={() => handleSelectRow(index)}
                        >
                          Usar neste cartao
                        </button>
                      </td>
                      <td>
                        <div className="table-actions">
                          {isEditing ? (
                            <>
                              <button
                                type="button"
                                className="table-button"
                                onClick={() => handleSaveEdit(index)}
                              >
                                Salvar
                              </button>
                              <button type="button" className="table-button" onClick={handleCancelEdit}>
                                Cancelar
                              </button>
                            </>
                          ) : (
                            <>
                              <button
                                type="button"
                                className="table-button"
                                onClick={() => handleStartEdit(index)}
                              >
                                Editar
                              </button>
                              <button
                                type="button"
                                className="table-button danger-button"
                                onClick={() => handleDeleteRow(index)}
                              >
                                Excluir
                              </button>
                            </>
                          )}
                        </div>
                      </td>
                    </tr>
                  )
                })}
              </tbody>
            </table>
          </div>
        ) : (
          <div className="empty-state">
            <h3>Pronto para receber a sua planilha</h3>
            <p>
              Assim que voce importar o Excel do perfil {currentProfile.label.toLowerCase()},
              vamos listar os aniversariantes e destacar quem faz aniversario hoje.
            </p>
          </div>
        )}
        </section>
      ) : null}
        </>
      ) : null}
    </main>
  )
}

export default App
