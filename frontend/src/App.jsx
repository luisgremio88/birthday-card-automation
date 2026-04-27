import { useEffect, useMemo, useState } from 'react'
import './App.css'

const REQUIRED_COLUMNS = ['nome', 'tabelionato', 'email', 'data_aniversario']
const API_BASE = `http://${window.location.hostname}/projeto-aniversario/backend`
const DEFAULT_PREVIEW_NAME = 'NOME DO ANIVERSARIANTE'
const SCREENSHOT_MODE = new URLSearchParams(window.location.search).get('shot')
const ASSET_BASE = import.meta.env.BASE_URL

const PROFILES = {
  associado: {
    id: 'associado',
    label: 'Associado',
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
    templatePath: `${ASSET_BASE}templates/cartao-diretoria.png`,
    templateHint: 'templates/cartao_base_limpo_diretoria.png',
    subject: 'Parabens ao membro da Diretoria',
    bodyTitle: 'Parabens ao membro da Diretoria',
    bodyFallback: 'Diretoria CNB/RS',
    mailDescription:
      'Fluxo especifico para a Diretoria, com lista propria, layout anual proprio e mensagem dedicada.',
    nameBox: { x: 430, y: 665, width: 1115, height: 145, align: 'center' },
  },
}

const SAMPLE_ROWS = {
  associado: [
    {
      nome: 'Paula da Silveira Delvalhas',
      tabelionato: 'Tabelionato Pozo - Sao Vicente do Sul',
      email: 'paula@email.com',
      data_aniversario: '27/04',
    },
    {
      nome: 'Jorge Alberto Pozo Camargo',
      tabelionato: 'Tabelionato Pozo - Sao Vicente do Sul',
      email: 'jorge@email.com',
      data_aniversario: '27/04',
    },
  ],
  diretoria: [
    {
      nome: 'Joao Ricardo Souza',
      tabelionato: 'Diretoria CNB/RS',
      email: 'joao.ricardo@email.com',
      data_aniversario: '27/04',
    },
    {
      nome: 'Mariana Costa Ferreira',
      tabelionato: 'Diretoria CNB/RS',
      email: 'mariana.ferreira@email.com',
      data_aniversario: '05/05',
    },
  ],
}

const SAMPLE_HISTORY = {
  associado: [
    {
      timestamp: '2026-04-27T08:00:00',
      data_referencia: '27/04/2026',
      nome: 'Paula da Silveira Delvalhas',
      email: 'paula@email.com',
      status: 'enviado',
      detalhes: 'Cartao enviado com sucesso pelo Outlook.',
      arquivo_cartao_url: null,
    },
    {
      timestamp: '2026-04-27T08:03:00',
      data_referencia: '27/04/2026',
      nome: 'Jorge Alberto Pozo Camargo',
      email: 'jorge@email.com',
      status: 'rascunho',
      detalhes: 'Rascunho aberto para conferencia.',
      arquivo_cartao_url: null,
    },
  ],
  diretoria: [
    {
      timestamp: '2026-04-27T08:10:00',
      data_referencia: '27/04/2026',
      nome: 'Joao Ricardo Souza',
      email: 'joao.ricardo@email.com',
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

  return String(value ?? '').trim()
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
  const [senderEmail, setSenderEmail] = useState('luis.dias@cnbrs.org.br')
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
    isScreenshotMode ? ['Paula da Silveira Delvalhas <paula@email.com>: enviado'] : [],
  )
  const [isRunningAutomation, setIsRunningAutomation] = useState(false)
  const [historyItems, setHistoryItems] = useState(
    isScreenshotMode ? SAMPLE_HISTORY.associado : [],
  )
  const [historyStatus, setHistoryStatus] = useState(
    isScreenshotMode ? '' : 'Carregando historico...',
  )

  const currentProfile = PROFILES[currentProfileId]

  const stats = useMemo(() => {
    const total = rows.length
    const uniqueUnits = new Set(rows.map((row) => row.tabelionato).filter(Boolean)).size
    const todayKey = new Intl.DateTimeFormat('pt-BR', {
      day: '2-digit',
      month: '2-digit',
      timeZone: 'America/Sao_Paulo',
    })
      .format(new Date())
      .replace('-', '/')
    const birthdaysToday = rows.filter(
      (row) => String(row.data_aniversario).slice(0, 5) === todayKey,
    ).length

    return { total, uniqueUnits, birthdaysToday, todayKey }
  }, [rows])

  const selectedRow = selectedRowIndex !== null ? rows[selectedRowIndex] ?? null : null
  const recipientEmail = selectedRow?.email || ''
  const emailSubject = currentProfile.subject
  const emailBodyTitle = selectedRow
    ? `${currentProfile.bodyTitle}, ${selectedRow.nome}!`
    : `${currentProfile.bodyTitle}!`
  const emailBodyUnit = selectedRow?.tabelionato || currentProfile.bodyFallback

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
    resetStatuses()
  }

  async function saveExcelOnServer(file) {
    const formData = new FormData()
    formData.append('excel', file)
    formData.append('profile', currentProfileId)

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
    setUploadStatus(`Salvando planilha do perfil ${currentProfile.label.toLowerCase()} no servidor...`)
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

        const formattedRows = rawRows
          .map((rawRow) => {
            const normalizedRow = Object.fromEntries(
              Object.entries(rawRow).map(([key, value]) => [normalizeHeader(key), value]),
            )

            return {
              nome: String(normalizedRow.nome ?? '').trim(),
              tabelionato: String(normalizedRow.tabelionato ?? '').trim(),
              email: String(normalizedRow.email ?? '').trim(),
              data_aniversario: excelDateToString(normalizedRow.data_aniversario, xlsx),
            }
          })
          .filter((row) => Object.values(row).some(Boolean))

        if (rawRows.length === 0) {
          setRows([])
          setSelectedRowIndex(null)
          setPreviewName(DEFAULT_PREVIEW_NAME)
          setError('A planilha esta vazia. Adicione os aniversariantes e tente novamente.')
          return
        }

        const originalHeaders = Object.keys(rawRows[0] ?? {}).map(normalizeHeader)
        const absentHeaders = REQUIRED_COLUMNS.filter((column) => !originalHeaders.includes(column))

        if (absentHeaders.length > 0) {
          setRows([])
          setSelectedRowIndex(null)
          setPreviewName(DEFAULT_PREVIEW_NAME)
          setError(`Colunas obrigatorias ausentes: ${absentHeaders.join(', ')}`)
          return
        }

        setRows(formattedRows)
        setSelectedRowIndex(formattedRows.length > 0 ? 0 : null)
        setPreviewName(formattedRows[0]?.nome || DEFAULT_PREVIEW_NAME)
        const uploadResult = await saveExcelOnServer(file)
        setUploadStatus(uploadResult.message || 'Planilha salva com sucesso.')
      } catch {
        setRows([])
        setSelectedRowIndex(null)
        setPreviewName(DEFAULT_PREVIEW_NAME)
        setUploadStatus('')
        setError('Nao foi possivel ler ou salvar o arquivo. Verifique se ele esta no formato .xlsx ou .xls.')
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

  async function runAutomation(mode) {
    if (isScreenshotMode) {
      setAutomationStatus(
        mode === 'send'
          ? 'Demo: envio automatico exibido apenas para captura do portfolio.'
          : 'Demo: rascunho do Outlook simulado para a captura do portfolio.',
      )
      setAutomationOutput(
        mode === 'send'
          ? ['Paula da Silveira Delvalhas <paula@email.com>: enviado']
          : ['Paula da Silveira Delvalhas <paula@email.com>: rascunho'],
      )
      return
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
        }),
      })

      const payload = await response.json()

      if (!response.ok || !payload.success) {
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
      setAutomationOutput([])
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
          <span className="eyebrow">Portfolio com 2 fluxos</span>
          <h1>Sistema de Cartoes de Aniversario</h1>
          <p className="lead">
            Gerencie no mesmo sistema os envios de <strong>Associado</strong> e
            <strong> Diretoria</strong>, cada um com planilha, template PSD anual e
            mensagem propria.
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
      </section>
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
              value={recipientEmail}
              readOnly
              placeholder="O e-mail vira automaticamente do Excel"
            />
          </div>

          <div className="mail-row">
            <div className="mail-tag">Cc</div>
            <input className="mail-input" type="text" value="" readOnly placeholder="Opcional para uma proxima etapa" />
          </div>

          <div className="mail-row">
            <div className="mail-tag">Cco</div>
            <input className="mail-input" type="text" value="" readOnly placeholder="Opcional para uma proxima etapa" />
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
            <strong> Para</strong> vem da planilha do perfil {currentProfile.label.toLowerCase()}.
          </p>
          <p>
            O assunto e o corpo acompanham o fluxo atual. Para a Diretoria, por exemplo,
            a mensagem muda para <code>Parabens ao membro da Diretoria</code>.
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
                      <span className={`status-chip status-${item.status}`}>{item.status}</span>
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
                  <th>Cartao</th>
                </tr>
              </thead>
              <tbody>
                {rows.map((row, index) => (
                  <tr
                    key={`${currentProfileId}-${row.email}-${index}`}
                    className={selectedRowIndex === index ? 'selected-row' : ''}
                  >
                    <td>{row.nome || 'Nao informado'}</td>
                    <td>{row.tabelionato || 'Nao informado'}</td>
                    <td>{row.email || 'Nao informado'}</td>
                    <td>{formatDateLabel(row.data_aniversario)}</td>
                    <td>
                      <button
                        type="button"
                        className="table-button"
                        onClick={() => handleSelectRow(index)}
                      >
                        Usar neste cartao
                      </button>
                    </td>
                  </tr>
                ))}
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
        </>
      ) : null}
    </main>
  )
}

export default App
