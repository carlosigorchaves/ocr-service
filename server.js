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
const { createClient } = require('@supabase/supabase-js')
const ws = require('ws')

const execAsync = promisify(exec)
const app       = express()
const upload    = multer({ dest: os.tmpdir() })

const PORT             = process.env.PORT || 3002
const AUTENTIQUE_TOKEN = process.env.AUTENTIQUE_API_TOKEN
const SANDBOX          = process.env.AUTENTIQUE_SANDBOX !== 'false'
const OCR_SECRET       = process.env.OCR_SECRET || ''
const SUPABASE_URL     = process.env.NEXT_PUBLIC_SUPABASE_URL
const SUPABASE_KEY     = process.env.SUPABASE_SERVICE_ROLE_KEY

function sb() {
  return createClient(SUPABASE_URL, SUPABASE_KEY, {
    global: { fetch },
    realtime: { transport: ws },
  })
}

const ORIGENS = (process.env.ALLOWED_ORIGINS || 'https://enviodocumentos.vercel.app').split(',').map(x => x.trim())

// CORS: o navegador envia o PDF direto para cá (sem passar pelo Vercel, que limita a 4,5 MB)
app.use((req, res, next) => {
  const origin = req.headers.origin
  if (origin && (ORIGENS.includes(origin) || ORIGENS.includes('*'))) {
    res.setHeader('Access-Control-Allow-Origin', origin)
    res.setHeader('Vary', 'Origin')
    res.setHeader('Access-Control-Allow-Methods', 'POST, GET, OPTIONS')
    res.setHeader('Access-Control-Allow-Headers', 'Authorization, Content-Type, x-ocr-secret')
    res.setHeader('Access-Control-Max-Age', '86400')
  }
  if (req.method === 'OPTIONS') return res.sendStatus(204)
  next()
})

// Autenticação: segredo (servidor) OU token de login do Supabase (navegador)
app.use(async (req, res, next) => {
  if (req.path === '/health') return next()
  if (OCR_SECRET && req.headers['x-ocr-secret'] === OCR_SECRET) return next()
  const auth = req.headers.authorization || ''
  const token = auth.startsWith('Bearer ') ? auth.slice(7) : ''
  if (token) {
    try {
      const { data, error } = await sb().auth.getUser(token)
      if (!error && data?.user) { req.user = data.user; return next() }
    } catch (e) { console.error('[auth]', e.message) }
  }
  return res.status(401).json({ erro: 'Não autorizado. Faça login novamente.' })
})

app.use(express.json())
app.get('/health', (req, res) => res.json({ ok: true, sandbox: SANDBOX }))

// Matrícula: remove espaços e zeros à esquerda para casar OCR x Excel/RM (ex.: 000123 = 123)
const normMat = m => String(m || '').trim().replace(/^0+/, '')

function mapaDeLista(rows) {
  const mapa = {}
  rows.forEach(r => {
    const chave = normMat(r.matricula)
    if (chave && r.email && r.email.includes('@')) {
      mapa[chave] = { nome: r.nome || '', email: r.email, cargo: '', matricula: r.matricula, centro_custo: r.centro_custo || '', situacao: String(r.situacao || 'ativo').toLowerCase() === 'demitido' ? 'demitido' : 'ativo' }
    }
  })
  return mapa
}

function lerExcel(buffer) {
  const wb    = XLSX.read(buffer, { type: 'buffer' })
  const ws    = wb.Sheets[wb.SheetNames[0]]
  const rows  = XLSX.utils.sheet_to_json(ws, { defval: '' })
  const mapa  = {}
  const normK = k => String(k).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim()
  rows.forEach(row => {
    const c = {}
    for (const [k, v] of Object.entries(row)) {
      const nk = normK(k).replace(/[\s-]+/g,'_')
      if (['nome','email','matricula','cpf','cargo','centro_custo','situacao'].includes(nk)) c[nk] = String(v).trim()
    }
    const chave = normMat(c.matricula) || c.cpf
    if (chave && c.email && c.email.includes('@')) {
      const sit = String(c.situacao || 'ativo').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'')
      mapa[chave] = { nome: c.nome || '', email: c.email, cargo: c.cargo || '', matricula: c.matricula || '', centro_custo: c.centro_custo || '', situacao: /demit|inativ|desat/.test(sit) ? 'demitido' : 'ativo' }
    }
  })
  return mapa
}

function httpRequest(url, options, body) {
  return new Promise((resolve, reject) => {
    const req = https.request(url, options, res => {
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

  const json = await httpRequest('https://api.autentique.com.br/v2/graphql', {
    method: 'POST',
    headers: { Authorization: `Bearer ${AUTENTIQUE_TOKEN}`, ...form.getHeaders() },
  }, form)

  if (json.errors) throw new Error(json.errors[0]?.message || JSON.stringify(json.errors))
  return json.data.createDocument
}

// ---------- IDENTIFICAÇÃO DE PÁGINAS (matrícula ou nome, em qualquer posição) ----------
const RE_ACENTOS = new RegExp('[' + String.fromCharCode(768) + '-' + String.fromCharCode(879) + ']', 'g')
const semAcento = s => String(s || '').normalize('NFD').replace(RE_ACENTOS, '').toUpperCase()
// número de 3 a 8 dígitos que não seja parte de data (19/10/2026), valor (1.500,00) ou CPF/telefone
const RE_NUM = /(?<![\d\/.,-])(\d{3,8})(?![\d\/]|[.,]\d|-\d)/g
const numerosDoTexto = t => [...t.matchAll(RE_NUM)].map(m => m[1])

function montarIndice(colaboradores) {
  const entradas = Object.entries(colaboradores).map(([k, c]) => ({
    k, toks: [...new Set(semAcento(c.nome).split(/[^A-Z0-9]+/).filter(x => x.length >= 2))],
  })).filter(e => e.toks.length >= 2)
  return {
    porMat: colaboradores,
    // pessoas cujo nome (todas as palavras) aparece na página
    achaNomes(t) {
      const palavras = new Set(t.split(/[^A-Z0-9]+/))
      const achados = entradas.filter(e => e.toks.every(x => palavras.has(x)))
      // descarta "ANA SOUSA" quando "ANA LIGIA SOUSA" também foi encontrado
      return achados
        .filter(a => !achados.some(b => b !== a && b.toks.length > a.toks.length && a.toks.every(x => b.toks.includes(x))))
        .map(e => e.k)
    },
  }
}

// Devolve { tipo: 'ok' | 'conferir' | null, key, motivo, matLida }
function identificarPagina(texto, indice) {
  const t = semAcento(texto).replace(/[ \t]+/g, ' ')
  const nomes = indice.achaNomes(t)

  // 1) número logo depois do rótulo Matrícula / Chapa
  let matLida = null
  const rot = /(MATRICULA|CHAPA|\bMAT\b\.?)/g
  let m
  while ((m = rot.exec(t))) {
    const ini = m.index + m[0].length
    const nums = numerosDoTexto(t.slice(ini, ini + 160))
    if (!nums.length) continue
    const k = normMat(nums[0])
    if (indice.porMat[k]) {
      if (nomes.length && !nomes.includes(k)) {
        return { tipo: 'conferir', key: k, motivo: 'A matrícula e o nome da página apontam para pessoas diferentes' }
      }
      return { tipo: 'ok', key: k }
    }
    matLida = matLida || nums[0]
  }

  // 2) números da página que são matrículas conhecidas
  const conhecidos = [...new Set(numerosDoTexto(t).map(normMat).filter(k => indice.porMat[k]))]
  if (conhecidos.length === 1) {
    const k = conhecidos[0]
    if (nomes.includes(k)) return { tipo: 'ok', key: k }
    return { tipo: 'conferir', key: k, motivo: 'Matrícula encontrada sem rótulo e sem o nome na página' }
  }
  // 3) só o nome
  if (nomes.length === 1) return { tipo: 'conferir', key: nomes[0], motivo: 'Identificado somente pelo nome' }
  if (conhecidos.length > 1 || nomes.length > 1) {
    return { tipo: 'conferir', key: null, motivo: 'Mais de uma pessoa possível nesta página' }
  }
  return { tipo: null, matLida }
}

async function textoDaPagina(pdfPath, pg) {
  try {
    const { stdout } = await execAsync(`pdftotext -f ${pg} -l ${pg} -layout "${pdfPath}" -`, { maxBuffer: 5 * 1024 * 1024 })
    return stdout
  } catch { return '' }
}

const resumoConferir = (lista, indice) => lista.map(c => ({
  paginas: c.paginas, motivo: c.motivo,
  sugestao: c.key && indice.porMat[c.key] ? { matricula: indice.porMat[c.key].matricula, nome: indice.porMat[c.key].nome } : null,
}))

app.post('/processar', upload.fields([
  { name: 'pdf',   maxCount: 1 },
  { name: 'excel', maxCount: 1 },
]), async (req, res) => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'ocr_'))
  try {
    if (!req.files?.pdf)   return res.status(400).json({ erro: 'PDF obrigatório.' })
    const fonteRM = req.body.fonte === 'rm'
    if (!fonteRM && !req.files?.excel) return res.status(400).json({ erro: 'Planilha Excel obrigatória.' })

    const nomeDoc  = req.body.nomeDocumento || 'Contracheque'
    const mensagem = req.body.mensagem || ''
    const pdfPath  = req.files.pdf[0].path

    const colaboradores = fonteRM
      ? mapaDeLista(await carregarFuncionarios())
      : lerExcel(fs.readFileSync(req.files.excel[0].path))
    console.log(`[${fonteRM ? 'rm' : 'excel'}] ${Object.keys(colaboradores).length} colaboradores`)

    const { stdout: infoOut } = await execAsync(`pdfinfo "${pdfPath}"`)
    const pagesMatch   = infoOut.match(/Pages:\s+(\d+)/)
    const totalPaginas = pagesMatch ? parseInt(pagesMatch[1]) : 0
    if (!totalPaginas) return res.status(400).json({ erro: 'Não foi possível ler o PDF.' })

    console.log(`[ocr] processando ${totalPaginas} páginas...`)
    const indice = montarIndice(colaboradores)
    const ocrPy = path.join(__dirname, 'ocr.py')
    const grupos = {}
    const conferir = []
    const naoIdentificadas = []
    let anterior = null
    let anexadas = 0

    for (let pg = 1; pg <= totalPaginas; pg++) {
      const imgPrefix = path.join(tmpDir, 'pg')
      await execAsync(`pdftoppm -jpeg -r 200 -f ${pg} -l ${pg} "${pdfPath}" "${imgPrefix}"`)
      const files = fs.readdirSync(tmpDir).filter(f => f.startsWith('pg') && f.endsWith('.jpg'))
      if (!files.length) continue
      const imgPath = path.join(tmpDir, files[0])
      let id = null, matLida = null, nomeLido = ''
      try {
        // 1) texto digital da página (PDF gerado por sistema)
        const txt = await textoDaPagina(pdfPath, pg)
        if (txt.replace(/\s/g, '').length >= 40) { id = identificarPagina(txt, indice); matLida = id.matLida || null }
        // 2) OCR da faixa do topo (contracheque e fichas)
        if (!id?.tipo) {
          const { stdout } = await execAsync(`python3 ${ocrPy} "${imgPath}"`)
          const out = stdout.trim()
          if (out.includes('||')) {
            const [mat, nome] = out.split('||')
            nomeLido = (nome || '').trim()
            if (indice.porMat[normMat(mat)]) id = { tipo: 'ok', key: normMat(mat) }
            else matLida = matLida || mat.trim()
          }
        }
        // 3) OCR da página inteira, procurando matrícula ou nome em qualquer posição
        if (!id?.tipo) {
          const { stdout } = await execAsync(`python3 ${ocrPy} "${imgPath}" --texto`, { maxBuffer: 5 * 1024 * 1024 })
          id = identificarPagina(stdout, indice)
          matLida = matLida || id.matLida || null
        }
      } catch (err) { console.error(`[ocr] erro página ${pg}:`, err.message) }
      try { fs.unlinkSync(imgPath) } catch {}

      if (id?.tipo === 'ok') {
        const g = grupos[id.key] || (grupos[id.key] = { matricula: id.key, nomeOCR: indice.porMat[id.key]?.nome || nomeLido, paginas: [] })
        g.paginas.push(pg); anterior = g
      } else if (id?.tipo === 'conferir') {
        const c = { key: id.key, motivo: id.motivo, paginas: [pg] }
        conferir.push(c); anterior = c
      } else if (matLida) {
        // matrícula lida, mas que não existe na lista: sai como "não encontrada"
        const k = normMat(matLida)
        const g = grupos[k] || (grupos[k] = { matricula: k, nomeOCR: nomeLido, paginas: [] })
        g.paginas.push(pg); anterior = g
      } else if (anterior) {
        // página sem identificação: continuação do documento anterior
        anterior.paginas.push(pg); anexadas++
      } else {
        naoIdentificadas.push(pg)
      }
    }

    const itens = Object.values(grupos)
    console.log(`[ocr] ${itens.length} identificados, ${conferir.length} para conferir, ${naoIdentificadas.length} sem identificação, ${anexadas} páginas anexadas`)
    if (!itens.length) {
      return res.status(400).json({
        erro: conferir.length
          ? `Nenhuma página foi identificada com segurança; ${conferir.length} ficaram para conferir.`
          : 'Nenhum colaborador identificado pelo OCR.',
        paraConferir: resumoConferir(conferir, indice), naoIdentificadas,
      })
    }

    const loteId     = 'lote_ocr_' + Date.now()
    const resultados = []

    for (const item of itens) {
      try {
        const dadosExcel = colaboradores[normMat(item.matricula)]
        if (!dadosExcel) {
          resultados.push({ matricula: item.matricula, nome: item.nomeOCR, ok: false, erro: 'Matrícula não encontrada na planilha' })
          continue
        }
        if (dadosExcel.situacao === 'demitido') {
          resultados.push({ matricula: item.matricula, nome: dadosExcel.nome || item.nomeOCR, ok: false, erro: 'Colaborador demitido — ignorado' })
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

        // Salva no Supabase para o acompanhamento funcionar
        if (SUPABASE_URL && SUPABASE_KEY) {
          const { data: inserted } = await sb().from('colaboradores').insert({
            lote_id:             loteId,
            nome:                dadosExcel.nome || item.nomeOCR,
            email:               dadosExcel.email,
            matricula:           dadosExcel.matricula || item.matricula,
            cargo:               dadosExcel.cargo || '',
            centro_custo:        dadosExcel.centro_custo || null,
            situacao:            dadosExcel.situacao || 'ativo',
            status:              'enviado',
            enviado_em:          new Date().toISOString(),
            document_id:         doc.id,
            signature_public_id: sig?.public_id || null,
            link_assinatura:     sig?.link?.short_link || null,
          }).select().single()
          if (inserted) console.log(`[supabase] salvo id=${inserted.id}`)
        }

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
      paraConferir: resumoConferir(conferir, indice),
      naoIdentificadas,
      paginasAnexadas: anexadas,
      enviados,
      semMatricula: resultados.filter(r => !r.ok && r.erro?.includes('não encontrada')).length,
      erros:        resultados.filter(r => !r.ok && !r.erro?.includes('não encontrada') && !r.erro?.includes('demitido')).length,
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

// ---------- ENVIO PADRÃO: 1 PDF igual para todos os colaboradores da planilha ----------
function lerExcelLista(buffer) {
  const wb   = XLSX.read(buffer, { type: 'buffer' })
  const sh   = wb.Sheets[wb.SheetNames[0]]
  const rows = XLSX.utils.sheet_to_json(sh, { defval: '' })
  const CAMPOS = ['nome','email','cpf','matricula','cargo','centro_custo','situacao']
  const normK  = k => String(k).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[\s-]+/g,'_').trim()
  return rows.map(row => {
    const c = {}, extras = {}
    for (const [k, v] of Object.entries(row)) {
      const nk = normK(k)
      if (CAMPOS.includes(nk)) c[nk] = String(v).trim()
      else if (String(v).trim() !== '') extras[k] = String(v).trim()
    }
    const sit = String(c.situacao || 'ativo').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    c.situacao = /demit|inativ|desat/.test(sit) ? 'demitido' : 'ativo'
    c.extras = extras
    return c
  })
}

app.post('/enviar-padrao', upload.fields([
  { name: 'pdf',   maxCount: 1 },
  { name: 'excel', maxCount: 1 },
]), async (req, res) => {
  try {
    if (!req.files?.pdf)   return res.status(400).json({ erro: 'PDF obrigatório.' })
    const fonteRM = req.body.fonte === 'rm'
    if (!fonteRM && !req.files?.excel) return res.status(400).json({ erro: 'Planilha Excel obrigatória.' })

    const nomeDoc  = req.body.nomeDocumento || 'Documento para assinatura'
    const mensagem = req.body.mensagem || ''
    const pdfNome  = req.files.pdf[0].originalname || 'documento.pdf'
    const pdfBuf   = fs.readFileSync(req.files.pdf[0].path)

    let todos
    if (fonteRM) {
      const filtro = String(req.body.centroCusto || '').trim()
      todos = (await carregarFuncionarios())
        .filter(r => !filtro || String(r.centro_custo || '') === filtro)
        .map(r => ({
          nome: r.nome || '', email: r.email || '', matricula: r.matricula || '', cpf: '', cargo: '',
          centro_custo: r.centro_custo || '', extras: {},
          situacao: String(r.situacao || 'ativo').toLowerCase() === 'demitido' ? 'demitido' : 'ativo',
        }))
    } else {
      todos = lerExcelLista(fs.readFileSync(req.files.excel[0].path))
    }
    const validos = todos.filter(c => c.email && c.email.includes('@') && c.nome && c.matricula)
    const ignoradas = todos.length - validos.length
    if (!validos.length) {
      return res.status(400).json({ erro: 'Nenhum colaborador válido na planilha. Colunas obrigatórias: nome, email e matricula.' })
    }

    const loteId = 'lote_' + Date.now()
    const resultados = []

    for (let i = 0; i < validos.length; i++) {
      const col = validos[i]
      const base = {
        lote_id: loteId, nome: col.nome, email: col.email, cpf: col.cpf || null,
        matricula: col.matricula, cargo: col.cargo || null, centro_custo: col.centro_custo || null,
        situacao: col.situacao, extras: Object.keys(col.extras).length ? col.extras : null,
      }
      if (col.situacao === 'demitido') {
        await sb().from('colaboradores').insert({ ...base, status: 'pendente' })
        resultados.push({ email: col.email, ok: false, erro: 'Colaborador demitido — ignorado' })
        continue
      }
      try {
        const doc = await criarDocumento({
          nome: `${nomeDoc} - ${col.nome}`, pdfBuf, pdfNome,
          email: col.email, nomeColaborador: col.nome, mensagem,
        })
        const sig = doc.signatures?.[0]
        await sb().from('colaboradores').insert({
          ...base, status: 'enviado', enviado_em: new Date().toISOString(),
          document_id: doc.id, signature_public_id: sig?.public_id || null,
          link_assinatura: sig?.link?.short_link || null,
        })
        resultados.push({ email: col.email, ok: true, documentId: doc.id })
        console.log(`[padrao ok] ${col.nome} <${col.email}>`)
      } catch (err) {
        console.error(`[padrao erro] ${col.email}: ${err.message}`)
        await sb().from('colaboradores').insert({ ...base, status: 'pendente' })
        resultados.push({ email: col.email, ok: false, erro: err.message })
      }
      if (i < validos.length - 1) await new Promise(r => setTimeout(r, 1100))
    }

    res.json({
      ok: true, loteId,
      total: validos.length,
      enviados: resultados.filter(r => r.ok).length,
      erros: resultados.filter(r => !r.ok && !r.erro?.includes('demitido')).length,
      demitidos: resultados.filter(r => r.erro?.includes('demitido')).length,
      ignoradas,
      sandbox: SANDBOX,
      resultados,
    })
  } catch (err) {
    console.error('[enviar-padrao]', err)
    res.status(500).json({ erro: err.message })
  } finally {
    try { fs.unlinkSync(req.files?.pdf?.[0]?.path) } catch {}
    try { fs.unlinkSync(req.files?.excel?.[0]?.path) } catch {}
  }
})

// ---------- INTEGRACAO TOTVS RM (webservice wsConsultaSQL) ----------
const RM_URL      = process.env.RM_URL
const RM_USER     = process.env.RM_USER
const RM_PASSWORD = process.env.RM_PASSWORD
const RM_COLIGADA = process.env.RM_COLIGADA || '1'
const RM_CONSULTA = process.env.RM_CONSULTA

function xmlUnescape(t) {
  return t.split('&lt;').join('<').split('&gt;').join('>').split('&quot;').join('"').split('&apos;').join("'").split('&amp;').join('&')
}

function parseResultado(xml) {
  const rows = []
  let pos = 0
  while (true) {
    const a = xml.indexOf('<Resultado>', pos)
    if (a < 0) break
    const b = xml.indexOf('</Resultado>', a)
    if (b < 0) break
    const chunk = xml.slice(a + 11, b)
    const row = {}
    let p = 0
    while (true) {
      const st = chunk.indexOf('<', p)
      if (st < 0) break
      const en = chunk.indexOf('>', st)
      if (en < 0) break
      const tag = chunk.slice(st + 1, en)
      if (tag.startsWith('/')) { p = en + 1; continue }
      if (tag.endsWith('/')) { row[tag.slice(0, -1).toLowerCase()] = ''; p = en + 1; continue }
      const close = chunk.indexOf('</' + tag + '>', en)
      if (close < 0) break
      row[tag.toLowerCase()] = xmlUnescape(chunk.slice(en + 1, close)).trim()
      p = close + tag.length + 3
    }
    rows.push(row)
    pos = b + 12
  }
  return rows
}

function consultarRM() {
  return new Promise((resolve, reject) => {
    if (!RM_URL || !RM_USER || !RM_PASSWORD || !RM_CONSULTA ||
        [RM_USER, RM_PASSWORD, RM_CONSULTA].includes('PREENCHER')) {
      return reject(new Error('Variáveis do RM não configuradas no Railway (RM_URL, RM_USER, RM_PASSWORD, RM_CONSULTA).'))
    }
    const envelope =
      '<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:tot="http://www.totvs.com/">' +
      '<soapenv:Header/><soapenv:Body><tot:RealizarConsultaSQL>' +
      `<tot:codSentenca>${RM_CONSULTA}</tot:codSentenca>` +
      `<tot:codColigada>${RM_COLIGADA}</tot:codColigada>` +
      '<tot:codSistema>P</tot:codSistema>' +
      `<tot:parameters>CODCOLIGADA=${RM_COLIGADA}</tot:parameters>` +
      '</tot:RealizarConsultaSQL></soapenv:Body></soapenv:Envelope>'
    const auth = Buffer.from(`${RM_USER}:${RM_PASSWORD}`).toString('base64')
    const req = https.request(RM_URL, {
      method: 'POST',
      timeout: 60000,
      headers: {
        'Content-Type': 'text/xml; charset=utf-8',
        'SOAPAction': 'http://www.totvs.com/IwsConsultaSQL/RealizarConsultaSQL',
        'Authorization': `Basic ${auth}`,
        'Content-Length': Buffer.byteLength(envelope),
      },
    }, res => {
      let data = ''
      res.setEncoding('utf8')
      res.on('data', c => data += c)
      res.on('end', () => {
        if (res.statusCode === 401) return reject(new Error('RM recusou o usuário ou a senha (HTTP 401).'))
        if (data.includes('faultstring') || data.includes(':Fault>')) {
          const i = data.indexOf('faultstring')
          const msg = i >= 0 ? data.slice(data.indexOf('>', i) + 1, data.indexOf('<', data.indexOf('>', i) + 1)) : 'erro SOAP'
          return reject(new Error('RM devolveu erro: ' + xmlUnescape(msg).slice(0, 300)))
        }
        const st = data.indexOf('<RealizarConsultaSQLResult')
        const en = data.indexOf('</RealizarConsultaSQLResult>')
        if (st < 0 || en < 0) return reject(new Error(`Resposta inesperada do RM (HTTP ${res.statusCode}).`))
        const inner = data.slice(data.indexOf('>', st) + 1, en)
        const xml = xmlUnescape(inner)
        const rows = parseResultado(xml)
        if (!rows.length && xml.trim() && !xml.includes('<NewDataSet')) {
          return reject(new Error('RM respondeu: ' + xml.trim().slice(0, 300)))
        }
        resolve(rows)
      })
    })
    req.on('timeout', () => req.destroy(new Error('Tempo esgotado ao falar com o RM (possível bloqueio de rede/geolocalização).')))
    req.on('error', e => reject(new Error('Falha ao conectar no RM: ' + e.message)))
    req.write(envelope)
    req.end()
  })
}

// Teste: devolve só contagem e nomes de colunas (nenhum dado pessoal)
app.get('/rm/testar', async (req, res) => {
  try {
    const rows = await consultarRM()
    const colunas = rows.length ? Object.keys(rows[0]) : []
    const esperadas = ['matricula', 'nome', 'email', 'centro_custo', 'situacao']
    const comEmail = rows.filter(r => r.email && r.email.includes('@')).length
    const demitidos = rows.filter(r => String(r.situacao || '').toLowerCase() === 'demitido').length
    res.json({
      ok: true, linhas: rows.length, colunas,
      faltando: esperadas.filter(c => !colunas.includes(c)),
      comEmailValido: comEmail, demitidos,
    })
  } catch (err) {
    console.error('[rm/testar]', err.message)
    res.status(502).json({ erro: err.message })
  }
})

// ── Cadastro local (tabela funcionarios): carga inicial + atualizações do RM ──
// Empresa dona do cadastro/RM desta instalação (outros clientes terão outro valor)
const EMPRESA = process.env.RM_EMPRESA || 'instituto-se-liga'
const normSit = s => String(s || 'ativo').toLowerCase() === 'demitido' ? 'demitido' : 'ativo'

async function lerCadastro() {
  const out = []
  for (let from = 0; ; from += 1000) {
    const { data, error } = await sb().from('funcionarios')
      .select('matricula,nome,email,centro_custo,situacao').eq('empresa', EMPRESA).order('matricula').range(from, from + 999)
    if (error) throw new Error('Cadastro: ' + error.message)
    out.push(...data)
    if (data.length < 1000) break
  }
  return out
}

// Usa o cadastro local; se ainda estiver vazio (sem carga inicial), consulta o RM direto.
async function carregarFuncionarios() {
  const cad = await lerCadastro()
  return cad.length ? cad : await consultarRM()
}

let sincronizando = false
async function sincronizarRM() {
  if (sincronizando) throw new Error('Já existe uma sincronização em andamento.')
  sincronizando = true
  try {
    const rm = await consultarRM()
    const novo = new Map()
    for (const r of rm) {
      const mat = normMat(r.matricula)
      if (!mat || !r.nome) continue
      novo.set(mat, {
        empresa: EMPRESA, matricula: mat, nome: String(r.nome).trim(), email: String(r.email || '').trim().toLowerCase(),
        centro_custo: String(r.centro_custo || '').trim(), situacao: normSit(r.situacao),
      })
    }
    if (!novo.size) throw new Error('O RM não retornou funcionários; cadastro mantido sem alterações.')

    const atual = new Map((await lerCadastro()).map(f => [f.matricula, f]))
    const mudou = (a, b) => a.nome !== b.nome || (a.email || '') !== b.email ||
      (a.centro_custo || '') !== b.centro_custo || a.situacao !== b.situacao

    const gravar = []
    let novos = 0, atualizados = 0, inalterados = 0
    for (const [mat, f] of novo) {
      const a = atual.get(mat)
      if (!a) { novos++; gravar.push(f) }
      else if (mudou(a, f)) { atualizados++; gravar.push(f) }
      else inalterados++
    }
    const agora = new Date().toISOString()
    for (let i = 0; i < gravar.length; i += 500) {
      const lote = gravar.slice(i, i + 500).map(f => ({ ...f, origem: 'rm', sincronizado_em: agora, updated_at: agora }))
      const { error } = await sb().from('funcionarios').upsert(lote, { onConflict: 'empresa,matricula' })
      if (error) throw new Error('Gravação: ' + error.message)
    }
    // quem sumiu do RM e ainda está ativo no cadastro é marcado como demitido (nada é apagado)
    const sumiu = [...atual.values()].filter(a => !novo.has(a.matricula) && a.situacao !== 'demitido').map(a => a.matricula)
    for (let i = 0; i < sumiu.length; i += 200) {
      const { error } = await sb().from('funcionarios')
        .update({ situacao: 'demitido', updated_at: agora, sincronizado_em: agora }).eq('empresa', EMPRESA).in('matricula', sumiu.slice(i, i + 200))
      if (error) throw new Error('Baixa: ' + error.message)
    }
    const resumo = { empresa: EMPRESA, total: novo.size, novos, atualizados, inalterados, removidos: sumiu.length }
    await sb().from('sincronizacoes_rm').insert(resumo)
    return resumo
  } catch (err) {
    try { await sb().from('sincronizacoes_rm').insert({ empresa: EMPRESA, erro: err.message }) } catch {}
    throw err
  } finally {
    sincronizando = false
  }
}

app.post('/rm/sincronizar', async (req, res) => {
  try {
    const resumo = await sincronizarRM()
    console.log('[rm/sincronizar]', JSON.stringify(resumo))
    res.json({ ok: true, ...resumo })
  } catch (err) {
    console.error('[rm/sincronizar]', err.message)
    res.status(502).json({ erro: err.message })
  }
})

app.get('/rm/ultima-sincronizacao', async (req, res) => {
  const { data } = await sb().from('sincronizacoes_rm').select('*').eq('empresa', EMPRESA).order('id', { ascending: false }).limit(1)
  res.json({ ok: true, ultima: data?.[0] || null })
})

// Lista normalizada para a tela "Buscar do RM" (lê do cadastro local)
app.get('/rm/colaboradores', async (req, res) => {
  try {
    const rows = await carregarFuncionarios()
    const lista = rows.map(r => ({
      nome: r.nome || '', email: r.email || '', matricula: r.matricula || '',
      centro_custo: r.centro_custo || '',
      situacao: String(r.situacao || 'ativo').toLowerCase() === 'demitido' ? 'demitido' : 'ativo',
    }))
    res.json({ ok: true, total: lista.length, colaboradores: lista })
  } catch (err) {
    console.error('[rm/colaboradores]', err.message)
    res.status(502).json({ erro: err.message })
  }
})

app.listen(PORT, () => console.log(`OCR Service na porta ${PORT} | Sandbox: ${SANDBOX}`))
