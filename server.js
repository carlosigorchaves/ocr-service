require('dotenv').config()
const express   = require('express')
const multer    = require('multer')
const { exec }  = require('child_process')
const { promisify } = require('util')
const fs        = require('fs')
const path      = require('path')
const os        = require('os')
const fetch     = require('node-fetch')
const FormData  = require('form-data')
const { createClient } = require('@supabase/supabase-js')

const execAsync = promisify(exec)
const app       = express()
const upload    = multer({ dest: os.tmpdir() })

const PORT             = process.env.PORT || 3002
const AUTENTIQUE_TOKEN = process.env.AUTENTIQUE_API_TOKEN
const SANDBOX          = process.env.AUTENTIQUE_SANDBOX !== 'false'
const SUPABASE_URL     = process.env.NEXT_PUBLIC_SUPABASE_URL
const SUPABASE_KEY     = process.env.SUPABASE_SERVICE_ROLE_KEY
const OCR_SECRET       = process.env.OCR_SECRET || ''

function sb() { return createClient(SUPABASE_URL, SUPABASE_KEY) }

// Autenticação por secret
app.use((req, res, next) => {
  if (req.path === '/health') return next()
  const secret = req.headers['x-ocr-secret']
  if (OCR_SECRET && secret !== OCR_SECRET) return res.status(401).json({ erro: 'Não autorizado.' })
  next()
})

app.use(express.json())

// Health check
app.get('/health', (req, res) => res.json({ ok: true, sandbox: SANDBOX }))

// Processar PDF com OCR
app.post('/processar', upload.single('pdf'), async (req, res) => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'ocr_'))

  try {
    if (!req.file) return res.status(400).json({ erro: 'PDF obrigatório.' })

    const nomeDoc  = req.body.nomeDocumento || 'Contracheque'
    const mensagem = req.body.mensagem || ''
    const pdfPath  = req.file.path

    // Conta páginas
    const { stdout: infoOut } = await execAsync(`pdfinfo "${pdfPath}"`)
    const pagesMatch   = infoOut.match(/Pages:\s+(\d+)/)
    const totalPaginas = pagesMatch ? parseInt(pagesMatch[1]) : 0
    if (!totalPaginas) return res.status(400).json({ erro: 'Não foi possível ler o PDF.' })

    console.log(`[ocr] processando ${totalPaginas} páginas...`)

    // Fase 1: OCR — agrupa páginas por colaborador
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
          const [matricula, nome] = output.split('||')
          const chave = `${matricula.trim()}|${nome.trim()}`
          if (!mapa[chave]) mapa[chave] = { matricula: matricula.trim(), nome: nome.trim(), paginas: [] }
          mapa[chave].paginas.push(pg)
        } else {
          console.log(`[ocr] página ${pg} não identificada`)
        }
      } catch (err) { console.error(`[ocr] erro página ${pg}:`, err.message) }
      try { fs.unlinkSync(imgPath) } catch {}
    }

    const colaboradores = Object.values(mapa)
    console.log(`[ocr] ${colaboradores.length} colaboradores identificados`)
    if (!colaboradores.length) return res.status(400).json({ erro: 'Nenhum colaborador identificado.' })

    // Fase 2: Separa PDFs e envia para Autentique
    const loteId     = 'lote_ocr_' + Date.now()
    const resultados = []

    for (const col of colaboradores) {
      try {
        // Separa PDF individual
        const paginasStr   = col.paginas.join(' ')
        const pdfIndivPath = path.join(tmpDir, `${col.matricula}.pdf`)
        await execAsync(`pdftk "${pdfPath}" cat ${paginasStr} output "${pdfIndivPath}"`)
        const pdfBuf = fs.readFileSync(pdfIndivPath)

        // Busca por matrícula no Supabase
        let encontrado = null
        const { data: r1 } = await sb().from('colaboradores')
          .select('id, email, nome').eq('matricula', col.matricula).limit(1)
        if (r1?.length) {
          encontrado = r1[0]
        } else {
          // fallback: busca pelo CPF
          const { data: r2 } = await sb().from('colaboradores')
            .select('id, email, nome').eq('cpf', col.matricula).limit(1)
          if (r2?.length) encontrado = r2[0]
        }

        const email    = encontrado?.email || null
        const nomeReal = encontrado?.nome   || col.nome

        if (!email) {
          console.log(`[aviso] matrícula ${col.matricula} sem email`)
          resultados.push({ matricula: col.matricula, nome: col.nome, ok: false, erro: 'Matrícula não encontrada no sistema' })
          continue
        }

        // Registra envio no Supabase
        const { data: inserted } = await sb().from('colaboradores').insert({
          lote_id:   loteId,
          nome:      nomeReal,
          email,
          matricula: col.matricula,
          status:    'pendente',
          extras:    { paginas: col.paginas },
        }).select().single()

        // Envia para Autentique
        const doc = await criarDocumento({
          nome:            `${nomeDoc} - ${nomeReal}`,
          pdfBuf,
          pdfNome:         `${col.matricula}.pdf`,
          email,
          nomeColaborador: nomeReal,
          mensagem,
        })

        const sig = doc.signatures?.[0]
        await sb().from('colaboradores').update({
          document_id:         doc.id,
          signature_public_id: sig?.public_id || null,
          link_assinatura:     sig?.link?.short_link || null,
          status:              'enviado',
          enviado_em:          new Date().toISOString(),
        }).eq('id', inserted.id)

        resultados.push({ matricula: col.matricula, nome: nomeReal, email, ok: true, documentId: doc.id })
        console.log(`[ok] ${nomeReal} <${email}> → ${doc.id}`)
        await new Promise(r => setTimeout(r, 1100))
      } catch (err) {
        console.error(`[erro] ${col.matricula}: ${err.message}`)
        resultados.push({ matricula: col.matricula, nome: col.nome, ok: false, erro: err.message })
      }
    }

    const enviados = resultados.filter(r => r.ok).length
    res.json({
      ok: true, loteId, totalPaginas,
      colaboradoresIdentificados: colaboradores.length,
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
    try { fs.unlinkSync(req.file?.path) } catch {}
  }
})

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

  const res  = await fetch('https://api.autentique.com.br/v2/graphql', {
    method:  'POST',
    headers: { Authorization: `Bearer ${AUTENTIQUE_TOKEN}`, ...form.getHeaders() },
    body:    form,
  })
  const json = await res.json()
  if (json.errors) throw new Error(json.errors[0]?.message || JSON.stringify(json.errors))
  return json.data.createDocument
}

app.listen(PORT, () => console.log(`OCR Service na porta ${PORT} | Sandbox: ${SANDBOX}`))
