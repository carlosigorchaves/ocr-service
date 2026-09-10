require('dotenv').config()
const express      = require('express')
const multer       = require('multer')
const { exec }     = require('child_process')
const { promisify }= require('util')
const fs           = require('fs')
const path         = require('path')
const os           = require('os')
const FormData     = require('form-data')
const XLSX         = require('xlsx')
const https        = require('https')
const http         = require('http') 

const execAsync = promisify(exec)
const app       = express()
const upload    = multer({ dest: os.tmpdir() })

const PORT             = process.env.PORT || 3002
const AUTENTIQUE_TOKEN = process.env.AUTENTIQUE_API_TOKEN
const SANDBOX          = process.env.AUTENTIQUE_SANDBOX !== 'false'
const OCR_SECRET       = process.env.OCR_SECRET || ''

app.use((req, res, next) => {
  if (req.path === '/health') return next()
  const secret = req.headers['x-ocr-secret']
  if (OCR_SECRET && secret !== OCR_SECRET) return res.status(401).json({ erro: 'Não autorizado.' })
  next()
})

app.use(express.json())
app.get('/health', (req, res) => res.json({ ok: true, sandbox: SANDBOX }))

function lerExcel(buffer) {
  const wb    = XLSX.read(buffer, { type: 'buffer' })
  const ws    = wb.Sheets[wb.SheetNames[0]]
  const rows  = XLSX.utils.sheet_to_json(ws, { defval: '' })
  const mapa  = {}
  const normK = k => String(k).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim()
  rows.forEach(row => {
    const c = {}
    for (const [k, v] of Object.entries(row)) {
      const nk = normK(k)
      if (['nome','email','matricula','cpf','cargo'].includes(nk)) c[nk] = String(v).trim()
    }
    const chave = c.matricula || c.cpf
    if (chave && c.email && c.email.includes('@')) {
      mapa[chave] = { nome: c.nome || '', email: c.email, cargo: c.cargo || '' }
    }
  })
  return mapa
}

// Requisição HTTP nativa (sem node-fetch)
function httpRequest(url, options, body) {
  return new Promise((resolve, reject) => {
    const lib     = url.startsWith('https') ? https : http
    const req     = lib.request(url, options, res => {
      let data = ''
      res.on('data', chunk => data += chunk)
      res.on('end', () => {
        try { resolve(JSON.parse(data)) }
        catch { reject(new Error('Resposta inválida: ' + data.slice(0, 200))) }
      })
    })
    req.on('error', reject)
    if (body) body.pipe(req)
    else req.end()
  })
}

async function criarDocumento({ nome, pdfBuf, pdfNome, email, nomeColaborador, mensagem }) {
  const query = `
    mutation CreateDocumentMutation($document: DocumentInput!, $signers: [SignerInput!]!, $file: Upload!) {
      createDocument(sandbox: ${SANDBOX}, document: $document, signers: $signers, file: $file) {
        id name created_at
        signatures { public_id name email action { name } link { short_link } }
      }
    }`

  const variables = {
    document: { name: nome, ...(mensagem ? { message: mensagem } : {}) },
    signers:  [{ email, name: nomeColaborador, action: 'SIGN' }],
    file: null,
  }

  const form = new FormData()
  form.append('operations', JSON.stringify({ query, variables }))
  form.append('map', JSON.stringify({ file: ['variables.file'] }))
  form.append('file', pdfBuf, { filename: pdfNome, contentType: 'application/pdf' })

  const headers = {
    Authorization: `Bearer ${AUTENTIQUE_TOKEN}`,
    ...form.getHeaders(),
  }

  const json = await httpRequest('https://api.autentique.com.br/v2/graphql', {
    method: 'POST',
    headers,
  }, form)

  if (json.errors) throw new Error(json.errors[0]?.message || JSON.stringify(json.errors))
  return json.data.createDocument
}

app.post('/processar', upload.fields([
  { name: 'pdf',   maxCount: 1 },
  { name: 'excel', maxCount: 1 },
]), async (req, res) => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'ocr_'))
  try {
    if (!req.files?.pdf)   return res.status(400).json({ erro: 'PDF obrigatório.' })
    if (!req.files?.excel) return res.status(400).json({ erro: 'Planilha Excel obrigatória.' })

    const nomeDoc  = req.body.nomeDocumento || 'Contracheque'
    const mensagem = req.body.mensagem || ''
    const pdfPath  = req.files.pdf[0].path

    const excelBuf      = fs.readFileSync(req.files.excel[0].path)
    const colaboradores = lerExcel(excelBuf)
    console.log(`[excel] ${Object.keys(colaboradores).length} colaboradores`)

    const { stdout: infoOut } = await execAsync(`pdfinfo "${pdfPath}"`)
    const pagesMatch   = infoOut.match(/Pages:\s+(\d+)/)
    const totalPaginas = pagesMatch ? parseInt(pagesMatch[1]) : 0
    if (!totalPaginas) return res.status(400).json({ erro: 'Não foi possível ler o PDF.' })

    console.log(`[ocr] processando ${totalPaginas} páginas...`)
    const mapa = {}

    for (let pg = 1; pg <= totalPaginas; pg++) {
      const imgPrefix = path.join(tmpDir, 'pg')
      await execAsync(`pdftoppm -jpeg -r 200 -f ${pg} -l ${pg} "${pdfPath}" "${imgPrefix}"`)
      const files = fs.readdirSync(tmpDir).filter(f => f.startsWith('pg') && f.endsWith('.jpg'))
      if (!files.length) continue
      const imgPath = path.join(tmpDir, files[0])
      try {
        const { stdout } = await execAsync(`python3 ${path.join(__dirname, 'ocr.py')} "${imgPath}"`)
        const output = stdout.trim()
        if (output && output !== 'NAO_IDENTIFICADO' && output.includes('||')) {
          const [matricula, nomeOCR] = output.split('||')
          const chave = matricula.trim()
          if (!mapa[chave]) mapa[chave] = { matricula: chave, nomeOCR: nomeOCR.trim(), paginas: [] }
          mapa[chave].paginas.push(pg)
        }
      } catch (err) { console.error(`[ocr] erro página ${pg}:`, err.message) }
      try { fs.unlinkSync(imgPath) } catch {}
    }

    const itens = Object.values(mapa)
    console.log(`[ocr] ${itens.length} identificados`)
    if (!itens.length) return res.status(400).json({ erro: 'Nenhum colaborador identificado pelo OCR.' })

    const loteId     = 'lote_ocr_' + Date.now()
    const resultados = []

    for (const item of itens) {
      try {
        const dadosExcel = colaboradores[item.matricula]
        if (!dadosExcel) {
          resultados.push({ matricula: item.matricula, nome: item.nomeOCR, ok: false, erro: 'Matrícula não encontrada na planilha' })
          continue
        }

        const paginasStr   = item.paginas.join(' ')
        const pdfIndivPath = path.join(tmpDir, `${item.matricula}.pdf`)
        await execAsync(`pdftk "${pdfPath}" cat ${paginasStr} output "${pdfIndivPath}"`)
        const pdfBuf = fs.readFileSync(pdfIndivPath)

        const doc = await criarDocumento({
          nome:            `${nomeDoc} - ${dadosExcel.nome || item.nomeOCR}`,
          pdfBuf,
          pdfNome:         `${item.matricula}.pdf`,
          email:           dadosExcel.email,
          nomeColaborador: dadosExcel.nome || item.nomeOCR,
          mensagem,
        })

        const sig = doc.signatures?.[0]
        resultados.push({
          matricula:      item.matricula,
          nome:           dadosExcel.nome,
          email:          dadosExcel.email,
          ok:             true,
          documentId:     doc.id,
          linkAssinatura: sig?.link?.short_link,
        })
        console.log(`[ok] ${dadosExcel.nome} <${dadosExcel.email}> → ${doc.id}`)
        await new Promise(r => setTimeout(r, 1100))
      } catch (err) {
        console.error(`[erro] ${item.matricula}: ${err.message}`)
        resultados.push({ matricula: item.matricula, nome: item.nomeOCR, ok: false, erro: err.message })
      }
    }

    const enviados = resultados.filter(r => r.ok).length
    res.json({
      ok: true, loteId, totalPaginas,
      colaboradoresIdentificados: itens.length,
      enviados,
      semMatricula: resultados.filter(r => !r.ok && r.erro?.includes('não encontrada')).length,
      erros:        resultados.filter(r => !r.ok && !r.erro?.includes('não encontrada')).length,
      sandbox:      SANDBOX,
      resultados,
    })
  } catch (err) {
    console.error('[processar]', err)
    res.status(500).json({ erro: err.message })
  } finally {
    try { fs.rmSync(tmpDir, { recursive: true }) } catch {}
    try { fs.unlinkSync(req.files?.pdf?.[0]?.path) } catch {}
    try { fs.unlinkSync(req.files?.excel?.[0]?.path) } catch {}
  }
})

app.listen(PORT, () => console.log(`OCR Service na porta ${PORT} | Sandbox: ${SANDBOX}`))
