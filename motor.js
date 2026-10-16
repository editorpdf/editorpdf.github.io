// Motor de PDF que roda dentro do navegador/celular, sem o servidor Python.
// Usa o MuPDF.js (o mesmo motor do PyMuPDF, compilado para WebAssembly), o pdf-lib (já carregado pela página)
// e o Tesseract.js para o OCR. Responde às mesmas rotas do servidor.py: /api/ping, /api/analyze, /api/apply,
// /api/ocr, /api/buscar e /api/tool (sem as ferramentas que dependem do computador: Office, Word/Excel/PowerPoint,
// HTML, PDF/A e certificado digital).
let M = null;                       // módulo MuPDF.js
const BASE = new URL('.', import.meta.url).href;
const { PDFDocument: LibDoc, StandardFonts, rgb, degrees } = window.PDFLib;

export async function iniciar() {
  if (!M) M = await import('./lib/mupdf/mupdf.js');
}

class Erro extends Error {}

/* ------------------------------------------------------------ utilitários */
const b64ParaBytes = b => { const s = atob(b), u = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) u[i] = s.charCodeAt(i); return u; };
const bytesParaB64 = u8 => { let s = ''; const CH = 0x8000; for (let i = 0; i < u8.length; i += CH) s += String.fromCharCode.apply(null, u8.subarray(i, i + CH)); return btoa(s); };
const base = nome => (String(nome || 'arquivo').replace(/\.[^.]*$/, '').replace(/[^\w\-. ()áéíóúâêôãõçÁÉÍÓÚÂÊÔÃÕÇ]/g, '_').slice(0, 80)) || 'arquivo';
const kb = n => n < 1048576 ? Math.max(1, Math.round(n / 1024)) + ' KB' : (n / 1048576).toFixed(1) + ' MB';
const MM = 72 / 25.4;
const num = (v, padrao, mini, maxi) => { let x = parseFloat(String(v ?? '').replace(',', '.')); if (!isFinite(x)) x = padrao; if (mini != null) x = Math.max(mini, x); if (maxi != null) x = Math.min(maxi, x); return x; };
const hexRgb = (h, p = [0, 0, 0]) => { h = String(h || '').replace('#', ''); if (h.length === 3) h = h.split('').map(c => c + c).join(''); const n = parseInt(h, 16); return isNaN(n) ? p : [(n >> 16 & 255) / 255, (n >> 8 & 255) / 255, (n & 255) / 255]; };
const corPdf = (h, p) => { const [r, g, b] = hexRgb(h, p); return rgb(r, g, b); };

function abrir(dados, nome = 'arquivo') {
  let doc;
  try { doc = M.Document.openDocument(dados, 'application/pdf'); } catch (e) { throw new Erro(`Não foi possível abrir “${nome}”: o PDF está corrompido ou não é um PDF.`); }
  if (doc.needsPassword()) throw new Erro(`“${nome}” tem senha. Use “Desbloquear PDF” primeiro.`);
  return doc.asPDF();
}
function salvar(doc, opcoes = 'garbage=compact,compress') { return doc.saveToBuffer(opcoes).asUint8Array().slice(); }
const rotacao = page => { try { const r = page.getObject().getInheritable('Rotate'); return r && r.isNumber() ? ((r.asNumber() % 360) + 360) % 360 : 0; } catch (e) { return 0; } };
// ponto/retângulo do espaço do MuPDF (origem no topo esquerdo) para o espaço do PDF (origem embaixo)
const ptPdf = (inv, x, y) => [x * inv[0] + y * inv[2] + inv[4], x * inv[1] + y * inv[3] + inv[5]];
function rectPdf(inv, r) { const a = ptPdf(inv, r[0], r[1]), b = ptPdf(inv, r[2], r[3]); return { x: Math.min(a[0], b[0]), y: Math.min(a[1], b[1]), w: Math.abs(b[0] - a[0]), h: Math.abs(b[1] - a[1]) }; }

/* '1-3,5,7-' -> índices */
function faixa(texto, total, vazioTodas = false) {
  texto = String(texto || '').trim().toLowerCase();
  if (!texto) { if (vazioTodas) return [...Array(total).keys()]; throw new Erro('Informe as páginas (por exemplo: 1-3, 5, 8-).'); }
  const out = [];
  for (const tok of texto.split(/[,;\s]+/)) {
    if (!tok) continue;
    let a, b, m;
    if ((m = tok.match(/^(\d+)?\s*-\s*(\d+)?$/)) && (m[1] || m[2])) { a = m[1] ? +m[1] : 1; b = m[2] ? +m[2] : total; }
    else if (/^\d+$/.test(tok)) a = b = +tok;
    else if (['fim', 'ultima', 'última'].includes(tok)) a = b = total;
    else throw new Erro(`Páginas inválidas: “${tok}”.`);
    if (a < 1 || b > total || a > b) throw new Erro(`As páginas ${tok} estão fora do documento (ele tem ${total}).`);
    for (let p = a; p <= b; p++) if (!out.includes(p - 1)) out.push(p - 1);
  }
  if (!out.length) throw new Erro('Nenhuma página escolhida.');
  return out;
}

/* ------------------------------------------------------------ fontes do pdf-lib */
const STD = { sans: ['Helvetica', 'HelveticaBold', 'HelveticaOblique', 'HelveticaBoldOblique'], serif: ['TimesRoman', 'TimesRomanBold', 'TimesRomanItalic', 'TimesRomanBoldItalic'], mono: ['Courier', 'CourierBold', 'CourierOblique', 'CourierBoldOblique'] };
async function fontePdf(pd, cache, fam = 'sans', bold = false, italic = false) {
  const k = (STD[fam] || STD.sans)[(bold ? 1 : 0) + (italic ? 2 : 0)];
  return cache[k] || (cache[k] = await pd.embedFont(StandardFonts[k]));
}
// troca os caracteres que a fonte padrão não tem por "?"
function cabe(fonte, txt, conta) {
  let out = '';
  for (const ch of txt) { try { fonte.encodeText(ch); out += ch; } catch (e) { out += '?'; if (conta) conta.n++; } }
  return out;
}

/* ------------------------------------------------------------ texto da página */
function corHex(c) {
  if (!c || !c.length) return '#000000';
  let r, g, b;
  if (c.length === 1) r = g = b = c[0];
  else if (c.length === 3) [r, g, b] = c;
  else { const [C, Mg, Y, K] = c; r = (1 - C) * (1 - K); g = (1 - Mg) * (1 - K); b = (1 - Y) * (1 - K); }
  return '#' + [r, g, b].map(v => Math.round(Math.max(0, Math.min(1, v)) * 255).toString(16).padStart(2, '0')).join('');
}
const quadRect = q => [Math.min(q[0], q[4]), Math.min(q[1], q[3]), Math.max(q[2], q[6]), Math.max(q[5], q[7])];
// blocos > linhas > caracteres (com fonte, tamanho, cor e posição de cada caractere)
function blocosDaPagina(page) {
  const st = page.toStructuredText('preserve-whitespace');
  const blocos = []; let b = null, l = null; const fontes = new Map();
  st.walk({
    beginTextBlock(bbox) { b = { bbox, lines: [] }; },
    beginLine(bbox, wmode, dir) { l = { bbox, dir, chars: [] }; },
    onChar(c, origin, font, size, quad, color) {
      let f = fontes.get(font.pointer);
      if (!f) { f = { name: font.getName(), bold: font.isBold(), italic: font.isItalic(), serif: font.isSerif(), mono: font.isMono() }; fontes.set(font.pointer, f); }
      l.chars.push({ c, origin, f, size, rect: quadRect(quad), color: corHex(color) });
    },
    endLine() { if (l && l.chars.length) b.lines.push(l); l = null; },
    endTextBlock() { if (b && b.lines.length) blocos.push(b); b = null; },
  });
  st.destroy();
  return blocos;
}
function familia(f) {
  const n = f.name || '';
  if (f.mono || /courier|mono|consolas|typewriter/i.test(n)) return 'mono';
  if ((f.serif || /times|garamond|georgia|palatino|cambria|book|minion|serif/i.test(n)) && !/sans|arial|helvet|calibri|verdana|tahoma/i.test(n)) return 'serif';
  return 'sans';
}
const ehNegrito = f => f.bold || /bold|black|heavy|semibold|demi/i.test(f.name || '');
const ehItalico = f => f.italic || /italic|oblique/i.test(f.name || '');
const semSubset = n => String(n || '').replace(/^[A-Z]{6}\+/, '');
const mediana = a => { const s = [...a].sort((x, y) => x - y), m = s.length >> 1; return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2; };

function analisar(dados, pno) {
  const doc = abrir(dados);
  try {
    if (!(pno >= 0 && pno < doc.countPages())) throw new Erro('página fora do documento');
    const page = doc.loadPage(pno);
    const [x0p, y0p, x1p, y1p] = page.getBounds();
    if (rotacao(page)) return { rotated: true, blocks: [] };
    const out = [];
    for (const b of blocosDaPagina(page)) {
      const lines = []; let horizontal = true;
      for (const l of b.lines) {
        if (Math.abs(l.dir[1]) > 0.02 || l.dir[0] < 0.98) { horizontal = false; break; }
        const text = l.chars.map(c => c.c).join('');
        if (!text.trim()) continue;
        // estilo principal: o que tem mais caracteres visíveis
        const cont = new Map();
        for (const c of l.chars) { if (!c.c.trim()) continue; const k = semSubset(c.f.name) + '|' + Math.round(c.size * 10) / 10 + '|' + c.color; cont.set(k, (cont.get(k) || 0) + 1); }
        let mainK = null, mx = -1; for (const [k, v] of cont) if (v > mx) { mx = v; mainK = k; }
        const main = l.chars.find(c => c.c.trim() && semSubset(c.f.name) + '|' + Math.round(c.size * 10) / 10 + '|' + c.color === mainK) || l.chars[0];
        const first = l.chars.find(c => c.c.trim()) || l.chars[0];
        lines.push({ bbox: l.bbox.map(v => Math.round(v * 100) / 100), origin: [Math.round(l.chars[0].origin[0] * 100) / 100, Math.round(first.origin[1] * 100) / 100],
          text, size: Math.round(main.size * 100) / 100, font: main.f.name, color: main.color, f: main.f, styles: [...cont.keys()] });
      }
      if (!horizontal || !lines.length) continue;
      const x0 = Math.min(...lines.map(l => l.bbox[0])), y0 = Math.min(...lines.map(l => l.bbox[1]));
      const x1 = Math.max(...lines.map(l => l.bbox[2])), y1 = Math.max(...lines.map(l => l.bbox[3]));
      const width = x1 - x0;
      const wrapped = lines.length > 1 && lines.slice(0, -1).every(l => (l.bbox[2] - l.bbox[0]) >= 0.85 * width);
      const text = wrapped ? lines.map((l, i) => i ? l.text.trim() : l.text.trimEnd()).join(' ') : lines.map(l => l.text.trimEnd()).join('\n');
      const size = mediana(lines.map(l => l.size));
      const deltas = lines.slice(0, -1).map((l, i) => lines[i + 1].origin[1] - l.origin[1]);
      let lineh = deltas.length && size ? mediana(deltas) / size : 1.2; lineh = Math.max(0.9, Math.min(2.5, lineh));
      let align = 'left';
      if (lines.length > 1) {
        const L = lines.map(l => l.bbox[0]), R = lines.map(l => l.bbox[2]), C = lines.map(l => (l.bbox[0] + l.bbox[2]) / 2);
        const sp = a => Math.max(...a) - Math.min(...a);
        align = sp(L) < 2 ? 'left' : sp(C) < 2 ? 'center' : sp(R) < 2 ? 'right' : 'left';
      }
      const f0 = lines[0].f, estilos = new Set(lines.flatMap(l => l.styles));
      out.push({ bbox: [x0, y0, x1, y1].map(v => Math.round(v * 100) / 100), lines: lines.map(({ f, styles, ...r }) => r), text, wrapped, size, font: f0.name,
        color: lines[0].color, family: familia(f0), bold: ehNegrito(f0), italic: ehItalico(f0), lineh: Math.round(lineh * 1000) / 1000, align, origin: lines[0].origin, mixed: estilos.size > 1 });
    }
    return { rotated: false, blocks: out, width: x1p - x0p, height: y1p - y0p };
  } finally { doc.destroy(); }
}

/* ------------------------------------------------------------ aplicar trocas de texto e tarjas */
function validar(ops, n) {
  const ok = v => typeof v === 'number' && isFinite(v) && Math.abs(v) < 1e6;
  if (!Array.isArray(ops) || ops.length > 2000) throw new Erro('lista de operações inválida');
  for (const o of ops) {
    if (!o || !['rep', 'redact'].includes(o.kind)) throw new Erro('operação desconhecida');
    if (!Number.isInteger(o.page) || o.page < 0 || o.page >= n) throw new Erro('página fora do documento');
    if (!Array.isArray(o.bbox) || o.bbox.length !== 4 || !o.bbox.every(ok)) throw new Erro('área inválida');
    if (o.kind === 'rep' && (!Array.isArray(o.origin) || !ok(o.size) || o.size < 1 || o.size > 500)) throw new Erro('troca de texto inválida');
  }
}
// quebra o texto em linhas que cabem na largura
function quebrar(fonte, texto, tam, larg) {
  const linhas = [];
  for (const par of String(texto).split('\n')) {
    const palavras = par.split(/(\s+)/); let atual = '';
    for (const p of palavras) {
      const t = atual + p;
      if (atual.trim() && fonte.widthOfTextAtSize(t.trimEnd(), tam) > larg) { linhas.push(atual.trimEnd()); atual = p.trimStart(); }
      else atual = t;
    }
    linhas.push(atual.trimEnd());
  }
  return linhas;
}
async function aplicar(dados, ops) {
  const doc = abrir(dados);
  validar(ops, doc.countPages());
  const avisos = [], report = [], invs = {};
  try {
    for (const pno of [...new Set(ops.map(o => o.page))]) {
      const page = doc.loadPage(pno);
      invs[pno] = M.Matrix.invert(page.getTransform());
      const reps = ops.filter(o => o.page === pno && o.kind === 'rep'), reds = ops.filter(o => o.page === pno && o.kind === 'redact');
      if (reps.length) {   // 1) apaga o texto antigo, sem pintar nada nem mexer em imagens e desenhos
        for (const o of reps) { const a = page.createAnnotation('Redact'); const [x0, y0, x1, y1] = o.bbox; a.setRect([x0 + 0.4, y0 + 0.4, x1 - 0.4, y1 - 0.4]); }
        page.applyRedactions(false, M.PDFPage.REDACT_IMAGE_NONE, M.PDFPage.REDACT_LINE_ART_NONE, M.PDFPage.REDACT_TEXT_REMOVE);
      }
      if (reds.length) {   // 2) tarjas: apagam texto, pixels de imagem e desenhos cobertos
        for (const o of reds) { const a = page.createAnnotation('Redact'); a.setRect(o.bbox); }
        page.applyRedactions(false, M.PDFPage.REDACT_IMAGE_PIXELS, M.PDFPage.REDACT_LINE_ART_REMOVE_IF_COVERED, M.PDFPage.REDACT_TEXT_REMOVE);
      }
    }
    dados = salvar(doc);
  } finally { doc.destroy(); }
  // 3) pinta as tarjas e escreve o texto novo (fontes padrão do PDF)
  const pd = await LibDoc.load(dados), pages = pd.getPages(), fontes = {}, conta = { n: 0 };
  for (const o of ops) {
    const pg = pages[o.page], inv = invs[o.page];
    if (o.kind === 'redact') { const r = rectPdf(inv, o.bbox); pg.drawRectangle({ x: r.x, y: r.y, width: r.w, height: r.h, color: corPdf(o.fill || '#000000') }); continue; }
    const texto = String(o.text || '');
    if (!texto.trim()) { report.push({ page: o.page, fonte: null, apagado: true }); continue; }
    const fonte = await fontePdf(pd, fontes, o.family || 'sans', !!o.bold, !!o.italic);
    report.push({ page: o.page, fonte: 'padrao' });
    const cor = corPdf(o.color), tam = +o.size, [x0, y0, x1, y1] = o.bbox;
    const multi = texto.includes('\n') || o.multi;
    if (!multi) {
      const t = cabe(fonte, texto, conta), w = fonte.widthOfTextAtSize(t, tam);
      let [ox, oy] = o.origin;
      if (o.align === 'center') ox = (x0 + x1) / 2 - w / 2; else if (o.align === 'right') ox = x1 - w;
      const p = ptPdf(inv, ox, oy); pg.drawText(t, { x: p[0], y: p[1], size: tam, font: fonte, color: cor });
    } else {
      const lineh = +(o.lineh || 1.2), larg = x1 - x0 + 2, pageBottom = pg.getSize().height;
      let s = tam, linhas, alt = (y1 - y0) + tam * 0.6, ok = false;
      for (let k = 0; k < 14; k++) {
        linhas = quebrar(fonte, cabe(fonte, texto, k ? null : conta), s, larg);
        if (linhas.length * s * lineh <= alt + s * 0.3) { ok = true; break; }
        if (y0 + alt < pageBottom - 4 && alt - (y1 - y0) < tam * 6) alt += tam; else s *= 0.94;
      }
      if (!ok) { avisos.push(`Página ${o.page + 1}: o texto não coube no bloco e foi reduzido ao máximo.`); }
      linhas.forEach((ln, i) => {
        if (!ln) return;
        const w = fonte.widthOfTextAtSize(ln, s);
        const lx = o.align === 'center' ? x0 + (larg - 2) / 2 - w / 2 : o.align === 'right' ? x1 - w : x0;
        const p = ptPdf(inv, lx, y0 + s * 0.93 + i * s * lineh);
        pg.drawText(ln, { x: p[0], y: p[1], size: s, font: fonte, color: cor });
      });
    }
  }
  if (conta.n) avisos.push(`${conta.n} caractere(s) não existem nas fontes padrão e viraram “?”.`);
  return { pdf: bytesParaB64(await pd.save()), warnings: avisos, report };
}

/* ------------------------------------------------------------ busca (tarja automática, buscar e substituir) */
const PADROES = {
  cpf: ['CPF', String.raw`(?<!\d)\d{3}\.?\d{3}\.?\d{3}\s?-?\s?\d{2}(?!\d)`],
  cnpj: ['CNPJ', String.raw`(?<!\d)\d{2}\.?\d{3}\.?\d{3}\s?\/?\s?\d{4}\s?-?\s?\d{2}(?!\d)`],
  rg: ['RG', String.raw`(?<![\d.])\d{1,2}\.\d{3}\.\d{3}-?[\dXx](?![\d.])`],
  email: ['E-mail', String.raw`[\w.+-]+@[\w-]+(?:\.[\w-]+)+`],
  telefone: ['Telefone', String.raw`(?<!\d)(?:\+?55\s?)?(?:\(?\d{2}\)?\s?)?9?\s?\d{4}\s?-?\s?\d{4}(?!\d)`],
  cep: ['CEP', String.raw`(?<!\d)\d{5}-\d{3}(?!\d)`],
  placa: ['Placa', String.raw`\b[A-Z]{3}-?\d[A-Z0-9]\d{2}\b`],
  data: ['Data', String.raw`(?<!\d)\d{2}\/\d{2}\/\d{4}(?!\d)`],
};
const escapar = s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
function expressoes(op) {
  const out = [];
  for (const k of op.padroes || []) if (PADROES[k]) out.push([PADROES[k][0], new RegExp(PADROES[k][1], 'gu')]);
  let termos = op.termos; if (typeof termos === 'string') termos = termos.split(/[\n;]+/);
  for (let t of termos || []) {
    t = String(t || '').trim(); if (!t) continue;
    if (t.length > 300) throw new Erro('Termo de busca grande demais.');
    let p = escapar(t); if (op.inteira) p = String.raw`(?<![\p{L}\p{N}_])` + p + String.raw`(?![\p{L}\p{N}_])`;
    out.push(['Termo', new RegExp(p, 'gu' + (op.maiusculas ? '' : 'i'))]);
  }
  if (!out.length) throw new Erro('Marque o que procurar ou digite uma palavra.');
  return out;
}
function ocorrencias(page, regs) {
  const achados = [];
  for (const b of blocosDaPagina(page)) for (const l of b.lines) {
    // o JS conta posições em unidades UTF-16: monta o texto e o mapa de posição -> caractere
    let texto = ''; const mapa = [];
    l.chars.forEach((c, i) => { for (let k = 0; k < c.c.length; k++) mapa.push(i); texto += c.c; });
    for (const [rot, rx] of regs) {
      rx.lastIndex = 0; let m;
      while ((m = rx.exec(texto))) {
        if (!m[0].length) { rx.lastIndex++; continue; }
        const a = mapa[m.index], z = mapa[m.index + m[0].length - 1];
        let r = null;
        for (let i = a; i <= z; i++) { const q = l.chars[i].rect; r = r ? [Math.min(r[0], q[0]), Math.min(r[1], q[1]), Math.max(r[2], q[2]), Math.max(r[3], q[3])] : [...q]; }
        if (r) achados.push([rot, m[0], r]);
      }
    }
  }
  const unicos = [];
  for (const a of achados) if (!unicos.some(u => Math.abs(a[2][0] - u[2][0]) < .5 && Math.abs(a[2][1] - u[2][1]) < .5 && Math.abs(a[2][2] - u[2][2]) < .5)) unicos.push(a);
  return unicos;
}
function buscar(dados, op) {
  const doc = abrir(dados), regs = expressoes(op), res = [];
  try {
    for (let i = 0; i < doc.countPages(); i++) {
      for (const [rotulo, texto, r] of ocorrencias(doc.loadPage(i), regs)) {
        res.push({ pagina: i, rotulo, texto, rect: r.map(v => Math.round(v * 100) / 100) });
        if (res.length >= 5000) return res;
      }
    }
  } finally { doc.destroy(); }
  return res;
}

/* ------------------------------------------------------------ OCR (Tesseract.js, sem internet) */
let tessCarregado = null;
function carregarTesseract() {
  return tessCarregado || (tessCarregado = new Promise((ok, falha) => {
    const s = document.createElement('script'); s.src = BASE + 'lib/tesseract/tesseract.min.js'; s.onload = () => ok(window.Tesseract); s.onerror = () => falha(new Erro('Não consegui carregar o OCR.')); document.head.appendChild(s);
  }));
}
async function ocr(dados, paginas, idioma = 'por+eng', pular = true) {
  const T = await carregarTesseract();
  const langs = String(idioma || 'por+eng').split('+').filter(l => ['por', 'eng'].includes(l));
  const worker = await T.createWorker(langs.length ? langs : ['por'], 1, { workerPath: BASE + 'lib/tesseract/worker.min.js', corePath: BASE + 'lib/tesseract/core', langPath: BASE + 'lib/tesseract/lang', gzip: false, workerBlobURL: false, cacheMethod: 'none' });
  const avisos = []; let reconhecidas = 0, giradas = 0;
  const doc = abrir(dados), pd = await LibDoc.load(dados), pages = pd.getPages();
  try {
    await worker.setParameters({ user_defined_dpi: '300' });
    const alvo = paginas && paginas.length ? paginas : [...Array(doc.countPages()).keys()];
    for (const i of alvo) {
      const page = doc.loadPage(i);
      if (pular) { const st = page.toStructuredText(''); const t = st.asText().trim(); st.destroy(); if (t.length > 20) continue; }
      if (rotacao(page)) { giradas++; continue; }
      const pix = page.toPixmap(M.Matrix.scale(300 / 72, 300 / 72), M.ColorSpace.DeviceRGB, false);
      const png = pix.asPNG().slice(); pix.destroy();
      const { data } = await worker.recognize(new Blob([png], { type: 'image/png' }), { pdfTitle: 'ocr', pdfTextOnly: true }, { pdf: true, text: false });
      // camada de texto invisível desenhada por cima da página original (a aparência não muda)
      const [emb] = await pd.embedPdf(new Uint8Array(data.pdf));
      const cb = pages[i].getCropBox();
      pages[i].drawPage(emb, { x: cb.x, y: cb.y, width: cb.width, height: cb.height });
      reconhecidas++;
    }
  } finally { await worker.terminate(); doc.destroy(); }
  if (giradas) avisos.push(`${giradas} página(s) girada(s) não foram reconhecidas no celular.`);
  return { pdf: bytesParaB64(await pd.save()), reconhecidas, avisos };
}

/* ------------------------------------------------------------ ZIP simples (vários arquivos de saída) */
const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
const crc32 = u => { let c = 0xFFFFFFFF; for (let i = 0; i < u.length; i++) c = CRC[(c ^ u[i]) & 255] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };
function zip(arquivos) {
  const enc = new TextEncoder(), partes = [], central = []; let pos = 0;
  for (const [nome, dados] of arquivos) {
    const n = enc.encode(nome), c = crc32(dados), h = new DataView(new ArrayBuffer(30));
    h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true); h.setUint32(14, c, true); h.setUint32(18, dados.length, true); h.setUint32(22, dados.length, true); h.setUint16(26, n.length, true);
    partes.push(new Uint8Array(h.buffer), n, dados);
    const d = new DataView(new ArrayBuffer(46));
    d.setUint32(0, 0x02014b50, true); d.setUint16(4, 20, true); d.setUint16(6, 20, true); d.setUint16(8, 0x0800, true); d.setUint32(16, c, true); d.setUint32(20, dados.length, true); d.setUint32(24, dados.length, true); d.setUint16(28, n.length, true); d.setUint32(42, pos, true);
    central.push(new Uint8Array(d.buffer), n);
    pos += 30 + n.length + dados.length;
  }
  const tam = central.reduce((a, p) => a + p.length, 0), e = new DataView(new ArrayBuffer(22));
  e.setUint32(0, 0x06054b50, true); e.setUint16(8, arquivos.length, true); e.setUint16(10, arquivos.length, true); e.setUint32(12, tam, true); e.setUint32(16, pos, true);
  const tudo = [...partes, ...central, new Uint8Array(e.buffer)], out = new Uint8Array(tudo.reduce((a, p) => a + p.length, 0));
  let k = 0; for (const p of tudo) { out.set(p, k); k += p.length; }
  return out;
}

/* ------------------------------------------------------------ imagens (canvas) */
async function abrirImg(nome, dados) {
  try { return await createImageBitmap(new Blob([dados]), { imageOrientation: 'from-image' }); }
  catch (e) { throw new Erro(`“${nome}” não é uma imagem válida (use JPG, PNG, WEBP, BMP ou GIF).`); }
}
async function canvasPara(cv, fmt, q = 0.9) {
  const mime = { jpg: 'image/jpeg', png: 'image/png', webp: 'image/webp' }[fmt] || 'image/png';
  if (mime === 'image/jpeg') { const c2 = document.createElement('canvas'); c2.width = cv.width; c2.height = cv.height; const x = c2.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, cv.width, cv.height); x.drawImage(cv, 0, 0); cv = c2; }
  const blob = await new Promise(r => cv.toBlob(r, mime, q));
  if (!blob || blob.type !== mime) throw new Erro(`Este aparelho não gera imagens ${fmt.toUpperCase()}.`);
  return new Uint8Array(await blob.arrayBuffer());
}
function desenhar(img, w, h, sx = 0, sy = 0, sw = img.width, sh = img.height, ang = 0, espelho = '') {
  const cv = document.createElement('canvas'), gira = ang === 90 || ang === 270;
  cv.width = gira ? h : w; cv.height = gira ? w : h;
  const x = cv.getContext('2d'); x.imageSmoothingQuality = 'high';
  x.translate(cv.width / 2, cv.height / 2); x.rotate(ang * Math.PI / 180);
  if (espelho === 'horizontal') x.scale(-1, 1); else if (espelho === 'vertical') x.scale(1, -1);
  x.drawImage(img, sx, sy, sw, sh, -w / 2, -h / 2, w, h);
  return cv;
}
const fmtDe = nome => { const e = (nome.split('.').pop() || '').toLowerCase(); return e === 'jpeg' ? 'jpg' : ['jpg', 'png', 'webp'].includes(e) ? e : 'png'; };

/* ------------------------------------------------------------ posição na página considerando a rotação (pdf-lib) */
// u,v: coordenadas como a página aparece na tela (origem embaixo à esquerda)
function noPapel(pg, u, v) {
  const cb = pg.getCropBox(), r = ((pg.getRotation().angle % 360) + 360) % 360, x0 = cb.x, y0 = cb.y, x1 = cb.x + cb.width, y1 = cb.y + cb.height;
  if (r === 90) return { x: x1 - v, y: y0 + u, rot: 90 };
  if (r === 180) return { x: x1 - u, y: y1 - v, rot: 180 };
  if (r === 270) return { x: x0 + v, y: y1 - u, rot: 270 };
  return { x: x0 + u, y: y0 + v, rot: 0 };
}
function tamanhoVisivel(pg) { const cb = pg.getCropBox(), r = ((pg.getRotation().angle % 360) + 360) % 360; return r === 90 || r === 270 ? [cb.height, cb.width] : [cb.width, cb.height]; }
function texto(pg, t, u, v, opt) { const p = noPapel(pg, u, v); pg.drawText(t, { ...opt, x: p.x, y: p.y, rotate: degrees(p.rot) }); }

/* ------------------------------------------------------------ ferramentas */
const F = {};
const PDF = ['.pdf'], IMG = ['.jpg', '.jpeg', '.png', '.webp', '.bmp', '.gif'];

F.juntar = { min: 2, max: 60, ext: PDF, fn(arqs) {
  const out = new M.PDFDocument();
  for (const [nome, dados] of arqs) { const d = abrir(dados, nome); for (let i = 0; i < d.countPages(); i++) out.graftPage(-1, d, i); d.destroy(); }
  const n = out.countPages(); const b = salvar(out); out.destroy();
  return [[['juntado.pdf', b]], [`${arqs.length} arquivos juntados, ${n} páginas no total.`]];
} };
function novoCom(src, pgs) { const o = new M.PDFDocument(); for (const p of pgs) o.graftPage(-1, src, p); const b = salvar(o); o.destroy(); return b; }
F.dividir = { min: 1, max: 10, ext: PDF, fn(arqs, op) {
  const saidas = [];
  for (const [nome, dados] of arqs) {
    const d = abrir(dados, nome), n = d.countPages(); let partes = [];
    if (op.modo === 'cada') { const k = Math.round(num(op.cada, 1, 1, 10000)); for (let i = 0; i < n; i += k) partes.push([...Array(Math.min(k, n - i)).keys()].map(j => i + j)); }
    else if (op.modo === 'paginas') partes = [...Array(n).keys()].map(i => [i]);
    else { for (const tok of String(op.intervalos || '').split(/[;,]+/)) if (tok.trim()) partes.push(faixa(tok, n)); if (!partes.length) throw new Erro('Informe os intervalos (por exemplo: 1-3, 4-6, 7-).'); }
    partes.forEach((pgs, i) => saidas.push([`${base(nome)}-parte-${String(i + 1).padStart(2, '0')}.pdf`, novoCom(d, pgs)]));
    d.destroy();
  }
  return [saidas, [`Gerados ${saidas.length} arquivos.`]];
} };
F.remover = { min: 1, max: 20, ext: PDF, fn(arqs, op) {
  return [arqs.map(([nome, dados]) => { const d = abrir(dados, nome), pgs = faixa(op.paginas, d.countPages());
    if (pgs.length >= d.countPages()) throw new Erro(`Isso apagaria todas as páginas de “${nome}”.`);
    pgs.sort((a, b) => b - a).forEach(p => d.deletePage(p)); const b = salvar(d); d.destroy(); return [base(nome) + '-sem-paginas.pdf', b]; }), []];
} };
F.extrair = { min: 1, max: 20, ext: PDF, fn(arqs, op) {
  return [arqs.map(([nome, dados]) => { const d = abrir(dados, nome), b = novoCom(d, faixa(op.paginas, d.countPages())); d.destroy(); return [base(nome) + '-extraido.pdf', b]; }), []];
} };
F.girar = { min: 1, max: 30, ext: PDF, fn(arqs, op) {
  const ang = Math.round(num(op.angulo, 90, 0, 270)); if (![90, 180, 270].includes(ang)) throw new Erro('Escolha 90, 180 ou 270 graus.');
  return [arqs.map(([nome, dados]) => { const d = abrir(dados, nome);
    for (const i of faixa(op.paginas, d.countPages(), true)) { const pg = d.loadPage(i); pg.getObject().put('Rotate', (rotacao(pg) + ang) % 360); }
    const b = salvar(d); d.destroy(); return [base(nome) + '-girado.pdf', b]; }), []];
} };
F.cortar = { min: 1, max: 20, ext: PDF, fn(arqs, op) {
  const auto = !!op.auto, [t, b, e, di] = ['topo', 'base', 'esq', 'dir'].map(k => num(op[k], 0, 0, 500) * MM);
  return [arqs.map(([nome, dados]) => { const d = abrir(dados, nome);
    for (let i = 0; i < d.countPages(); i++) {
      const pg = d.loadPage(i), inv = M.Matrix.invert(pg.getTransform()), [X0, Y0, X1, Y1] = pg.getBounds();
      let alvo;
      if (auto) {
        let r = null; const junta = q => { r = r ? [Math.min(r[0], q[0]), Math.min(r[1], q[1]), Math.max(r[2], q[2]), Math.max(r[3], q[3])] : [...q]; };
        const st = pg.toStructuredText('preserve-images,collect-vectors');
        st.walk({ beginTextBlock: junta, onImageBlock: junta, onVector(bb) { if (bb[2] - bb[0] < (X1 - X0) * .98 || bb[3] - bb[1] < (Y1 - Y0) * .98) junta(bb); } }); st.destroy();
        if (!r) continue;
        alvo = [Math.max(X0, r[0] - 14), Math.max(Y0, r[1] - 14), Math.min(X1, r[2] + 14), Math.min(Y1, r[3] + 14)];
      } else alvo = [X0 + e, Y0 + t, X1 - di, Y1 - b];
      if (alvo[2] - alvo[0] < 20 || alvo[3] - alvo[1] < 20) throw new Erro(`As margens são grandes demais para a página ${i + 1}.`);
      const rp = rectPdf(inv, alvo);
      pg.getObject().put('CropBox', [rp.x, rp.y, rp.x + rp.w, rp.y + rp.h]);
    }
    const out = salvar(d); d.destroy(); return [base(nome) + '-cortado.pdf', out]; }),
    ['As páginas foram cortadas (o conteúdo fora da área continua no arquivo, só fica escondido).']];
} };
F.comprimir = { min: 1, max: 20, ext: PDF, async fn(arqs, op) {
  const [lado, q] = { baixa: [2400, .85], media: [1700, .7], alta: [1100, .5] }[op.nivel] || [1700, .7];
  const saidas = [], avisos = [];
  for (const [nome, dados] of arqs) {
    const d = abrir(dados, nome);
    // refaz as fotos grandes em JPEG menor (as com transparência ficam como estão)
    for (let n = 1; n < d.countObjects(); n++) {
      let ref; try { ref = d.newIndirect(n); if (!ref.isStream()) continue; } catch (e) { continue; }
      const obj = ref.resolve();
      if (obj.get('Subtype').toString() !== '/Image' || obj.get('SMask').isNull() === false || obj.get('ImageMask').valueOf() === true || !obj.get('Mask').isNull()) continue;
      const w = obj.get('Width').valueOf(), h = obj.get('Height').valueOf(); if (!(w * h > 250000)) continue;
      try {
        const im = d.loadImage(ref), pix = im.toPixmap(), rgbPix = pix.getColorSpace() && pix.getColorSpace().isRGB() ? pix : pix.convertToColorSpace(M.ColorSpace.DeviceRGB, false);
        const bm = await createImageBitmap(new Blob([rgbPix.asPNG().slice()]));
        const esc = Math.min(1, lado / Math.max(w, h)), nw = Math.max(1, Math.round(w * esc)), nh = Math.max(1, Math.round(h * esc));
        const jpg = await canvasPara(desenhar(bm, nw, nh), 'jpg', q);
        if (jpg.length < ref.readRawStream().getLength() * 0.9) {
          ref.writeRawStream(jpg);
          ['DecodeParms', 'Decode', 'SMaskInData'].forEach(k => obj.delete(k));
          obj.put('Filter', d.newName('DCTDecode')); obj.put('Width', nw); obj.put('Height', nh); obj.put('BitsPerComponent', 8); obj.put('ColorSpace', d.newName('DeviceRGB'));
        }
      } catch (e) { /* imagem em formato que o navegador não decodifica: mantém */ }
    }
    try { d.subsetFonts(); } catch (e) {}
    let novo = salvar(d, 'garbage=deduplicate,compress,compress-fonts,compress-images,clean'); d.destroy();
    if (novo.length >= dados.length) { novo = dados; avisos.push(`“${nome}” já estava bem otimizado: mantive o original.`); }
    else avisos.push(`“${nome}”: de ${kb(dados.length)} para ${kb(novo.length)} (${Math.round(100 - 100 * novo.length / dados.length)}% menor).`);
    saidas.push([base(nome) + '-comprimido.pdf', novo]);
  }
  return [saidas, avisos];
} };
F.reparar = { min: 1, max: 20, ext: PDF, fn(arqs) {
  const s = [], a = [];
  for (const [nome, dados] of arqs) { const d = abrir(dados, nome); if (!d.countPages()) throw new Erro(`“${nome}” não tem nenhuma página recuperável.`);
    s.push([base(nome) + '-reparado.pdf', salvar(d, 'garbage=compact,compress,clean')]); a.push(`“${nome}”: ${d.wasRepaired() ? 'estrutura reparada' : 'não havia defeito aparente, reescrito limpo'} (${d.countPages()} páginas).`); d.destroy(); }
  return [s, a];
} };
F.ocr = { min: 1, max: 10, ext: PDF, async fn(arqs, op) {
  const s = [], a = [];
  for (const [nome, dados] of arqs) { const d = abrir(dados, nome), pgs = faixa(op.paginas, d.countPages(), true); d.destroy();
    const r = await ocr(dados, pgs, op.idioma || 'por+eng', op.pular !== false); s.push([base(nome) + '-ocr.pdf', b64ParaBytes(r.pdf)]);
    a.push(`“${nome}”: texto reconhecido em ${r.reconhecidas} página(s). ${r.avisos.join(' ')}`); }
  return [s, a];
} };
F.marca = { min: 1, max: 20, ext: PDF, async fn(arqs, op) {
  const opac = num(op.opacidade, 30, 1, 100) / 100, ang = num(op.angulo, 45, -180, 180), rep = op.posicao === 'repetir', cor = corPdf(op.cor, [.6, .6, .6]);
  const s = [];
  for (const [nome, dados] of arqs) {
    const pd = await LibDoc.load(dados); const fontes = {};
    let img = null, txt = null, fonte = null, tam = num(op.tamanho, 48, 6, 400);
    if (op.tipo === 'imagem') { if (!op.imagem || !op.imagem.b64) throw new Erro("Escolha a imagem da marca d'água."); const b = b64ParaBytes(op.imagem.b64); img = /\.png$/i.test(op.imagem.nome) ? await pd.embedPng(b) : await pd.embedJpg(b).catch(() => pd.embedPng(b)); }
    else { txt = String(op.texto || '').trim(); if (!txt) throw new Erro("Digite o texto da marca d'água."); fonte = await fontePdf(pd, fontes, 'sans', true); txt = cabe(fonte, txt); }
    const pages = pd.getPages();
    for (const i of faixa(op.paginas, pages.length, true)) {
      const pg = pages[i], [W, H] = tamanhoVisivel(pg);
      const alvos = !rep ? [[W / 2, H / 2]] : img ? [W * .25, W * .75].flatMap(x => [H * .2, H * .5, H * .8].map(y => [x, y])) : [W * .28, W * .72].flatMap(x => [H * .18, H * .38, H * .58, H * .78].map(y => [x, y]));
      const rad = ang * Math.PI / 180;
      for (const [cx, cy] of alvos) {
        let w, h; if (img) { const esc = num(op.escala, 50, 5, 300) / 100; w = img.width * esc * 72 / 96; h = img.height * esc * 72 / 96; if (w > W) { h *= W / w; w = W; } } else { w = fonte.widthOfTextAtSize(txt, tam); h = tam * .7; }
        // canto inferior esquerdo de um retângulo w x h girado em volta do centro
        const u = cx - (w / 2) * Math.cos(rad) + (h / 2) * Math.sin(rad), v = cy - (w / 2) * Math.sin(rad) - (h / 2) * Math.cos(rad);
        const p = noPapel(pg, u, v);
        if (img) pg.drawImage(img, { x: p.x, y: p.y, width: w, height: h, rotate: degrees(ang + p.rot), opacity: opac });
        else pg.drawText(txt, { x: p.x, y: p.y, size: tam, font: fonte, color: cor, rotate: degrees(ang + p.rot), opacity: opac });
      }
    }
    s.push([base(nome) + '-marca.pdf', await pd.save()]);
  }
  return [s, []];
} };
F.numerar = { min: 1, max: 20, ext: PDF, async fn(arqs, op) {
  const pos = op.posicao || 'baixo-centro', fmt = String(op.formato || '{n}').trim() || '{n}', ini = Math.round(num(op.inicio, 1, 0, 100000)), tam = num(op.tamanho, 11, 5, 60), marg = num(op.margem, 14, 4, 120) * MM, pular = !!op.pular_primeira;
  const s = [];
  for (const [nome, dados] of arqs) {
    const pd = await LibDoc.load(dados), fonte = await fontePdf(pd, {}, 'sans'), pages = pd.getPages(), total = pages.length;
    for (const i of faixa(op.paginas, total, true)) {
      if (pular && i === 0) continue;
      const pg = pages[i], [W, H] = tamanhoVisivel(pg), n = ini + (i - (pular ? 1 : 0));
      const t = cabe(fonte, fmt.replace(/\{n\}/g, n).replace(/\{total\}/g, total + ini - 1 - (pular ? 1 : 0))), w = fonte.widthOfTextAtSize(t, tam);
      const v = pos.startsWith('cima') ? H - marg - tam : marg, u = pos.endsWith('esquerda') ? marg : pos.endsWith('direita') ? W - marg - w : W / 2 - w / 2;
      texto(pg, t, u, v, { size: tam, font: fonte, color: corPdf(op.cor, [.2, .2, .2]) });
    }
    s.push([base(nome) + '-numerado.pdf', await pd.save()]);
  }
  return [s, []];
} };
F.cabecalho = { min: 1, max: 20, ext: PDF, async fn(arqs, op) {
  const campos = {}; for (const k of ['cab_esq', 'cab_centro', 'cab_dir', 'rod_esq', 'rod_centro', 'rod_dir']) campos[k] = String(op[k] || '').trim();
  const logoB = op.logo && op.logo.b64 ? b64ParaBytes(op.logo.b64) : null;
  if (!Object.values(campos).some(Boolean) && !logoB) throw new Erro('Escreva algum texto para o cabeçalho ou o rodapé (ou escolha um logotipo).');
  const tam = num(op.tamanho, 9, 5, 40), marg = num(op.margem, 10, 3, 60) * MM, cor = corPdf(op.cor, [.3, .3, .3]), hoje = new Date().toLocaleDateString('pt-BR');
  const lh = num(op.logo_altura, 12, 4, 60) * MM, lpos = op.logo_pos || 'cab_esq';
  const s = [];
  for (const [nome, dados] of arqs) {
    const pd = await LibDoc.load(dados), fonte = await fontePdf(pd, {}, 'sans'), pages = pd.getPages(), total = pages.length;
    let logo = null; if (logoB) { try { logo = await pd.embedPng(logoB); } catch (e) { try { logo = await pd.embedJpg(logoB); } catch (e2) { throw new Erro('O logotipo não é uma imagem válida (use PNG ou JPG).'); } } }
    for (const i of faixa(op.paginas, total, true)) {
      if (op.pular_primeira && i === 0) continue;
      const pg = pages[i], [W, H] = tamanhoVisivel(pg);
      for (const [k, t0] of Object.entries(campos)) {
        if (!t0) continue;
        const t = cabe(fonte, t0.replace(/\{n\}/g, i + 1).replace(/\{total\}/g, total).replace(/\{data\}/g, hoje).replace(/\{arquivo\}/g, base(nome))), w = fonte.widthOfTextAtSize(t, tam);
        const v = k.startsWith('cab') ? H - marg - tam : marg, u = k.endsWith('esq') ? marg : k.endsWith('dir') ? W - marg - w : W / 2 - w / 2;
        texto(pg, t, u, v, { size: tam, font: fonte, color: cor });
      }
      if (op.linha) {
        const ys = []; if (['cab_esq', 'cab_centro', 'cab_dir'].some(k => campos[k]) || (logo && lpos.startsWith('cab'))) ys.push(H - marg - tam * 1.6); if (['rod_esq', 'rod_centro', 'rod_dir'].some(k => campos[k])) ys.push(marg + tam * 1.4);
        for (const y of ys) { const a = noPapel(pg, marg, y), b = noPapel(pg, W - marg, y); pg.drawLine({ start: { x: a.x, y: a.y }, end: { x: b.x, y: b.y }, thickness: .5, color: cor }); }
      }
      if (logo) {
        const h = lh, w = h * logo.width / logo.height, v = lpos.startsWith('cab') ? H - marg - h + 2 : marg - 2, u = lpos.endsWith('esq') ? marg : lpos.endsWith('dir') ? W - marg - w : W / 2 - w / 2;
        const p = noPapel(pg, u, v); pg.drawImage(logo, { x: p.x, y: p.y, width: w, height: h, rotate: degrees(p.rot) });
      }
    }
    s.push([base(nome) + '-cabecalho.pdf', await pd.save()]);
  }
  return [s, []];
} };
F.proteger = { min: 1, max: 20, ext: PDF, fn(arqs, op) {
  const senha = String(op.senha || '');
  if (senha.length < 4) throw new Erro('Use uma senha com pelo menos 4 caracteres.');
  if (/[,=\s]/.test(senha)) throw new Erro('No celular, a senha não pode ter vírgula, sinal de igual nem espaço.');
  let p = -3904 | 512; if (op.imprimir !== false) p |= 4 | 2048; if (op.copiar !== false) p |= 16; if (op.editar) p |= 8 | 32 | 256 | 1024;
  return [arqs.map(([nome, dados]) => { const d = abrir(dados, nome); const b = salvar(d, `garbage=compact,compress,encrypt=aes-256,user-password=${senha},owner-password=${senha},permissions=${p}`); d.destroy(); return [base(nome) + '-protegido.pdf', b]; }),
    ['Para abrir, será pedida a senha que você digitou. Guarde-a: não há como recuperar.']];
} };
F.desbloquear = { min: 1, max: 20, ext: PDF, fn(arqs, op) {
  const a = [];
  const s = arqs.map(([nome, dados]) => {
    let doc; try { doc = M.Document.openDocument(dados, 'application/pdf'); } catch (e) { throw new Erro(`Não foi possível abrir “${nome}”.`); }
    if (doc.needsPassword()) { if (!op.senha || !doc.authenticatePassword(String(op.senha))) throw new Erro(`A senha não abre “${nome}”. Esta ferramenta só remove a proteção de quem sabe a senha.`); a.push(`“${nome}”: senha aceita, proteção removida.`); }
    else a.push(`“${nome}” abre sem senha: removi as restrições de impressão, cópia e edição.`);
    const b = salvar(doc.asPDF(), 'garbage=compact,compress,encrypt=none'); doc.destroy(); return [base(nome) + '-desbloqueado.pdf', b];
  });
  return [s, a];
} };
function linhasTexto(doc, i) { if (i >= doc.countPages()) return []; const st = doc.loadPage(i).toStructuredText(''); const t = st.asText(); st.destroy(); return t.split('\n').map(l => l.trim()).filter(Boolean); }
function diferencas(a, b) {   // diff de linhas (maior subsequência comum)
  const n = a.length, m = b.length, L = Array.from({ length: n + 1 }, () => new Uint32Array(m + 1));
  for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) L[i][j] = a[i] === b[j] ? L[i + 1][j + 1] + 1 : Math.max(L[i + 1][j], L[i][j + 1]);
  const out = []; let i = 0, j = 0;
  while (i < n && j < m) { if (a[i] === b[j]) { i++; j++; } else if (L[i + 1][j] >= L[i][j + 1]) out.push(['-', a[i++]]); else out.push(['+', b[j++]]); }
  while (i < n) out.push(['-', a[i++]]); while (j < m) out.push(['+', b[j++]]);
  return out;
}
F.comparar = { min: 2, max: 2, ext: PDF, async fn(arqs) {
  const [[na, da], [nb, db]] = arqs, A = abrir(da, na), B = abrir(db, nb);
  const pd = await LibDoc.create(), fontes = {}, f = await fontePdf(pd, fontes, 'sans'), fb = await fontePdf(pd, fontes, 'sans', true);
  let pg = null, y = 0;
  const linha = (t, cor = [0, 0, 0], tam = 9, negrito = false) => {
    const fo = negrito ? fb : f; t = cabe(fo, t);
    for (const parte of quebrar(fo, t, tam, 515)) { if (!pg || y < 42) { pg = pd.addPage([595, 842]); y = 792; } pg.drawText(parte, { x: 40, y, size: tam, font: fo, color: rgb(...cor) }); y -= tam * 1.5; }
  };
  linha('Comparação de PDFs', undefined, 16, true); linha(`A (antigo): ${na}  [${A.countPages()} páginas]`); linha(`B (novo):   ${nb}  [${B.countPages()} páginas]`);
  linha('Em vermelho: o que só existe em A. Em verde: o que só existe em B.', [.35, .35, .35]); y -= 8;
  let com = 0;
  for (let i = 0; i < Math.max(A.countPages(), B.countPages()); i++) {
    const d = diferencas(linhasTexto(A, i), linhasTexto(B, i)); if (!d.length) continue; com++;
    linha(`Página ${i + 1}`, undefined, 11, true);
    for (const [s, t] of d) linha((s === '-' ? '- ' : '+ ') + t, s === '-' ? [.75, .1, .1] : [.05, .5, .2]);
    y -= 6;
  }
  if (!com) linha('Nenhuma diferença de texto encontrada.', [.05, .5, .2], 12, true);
  A.destroy(); B.destroy();
  return [[['comparacao.pdf', await pd.save()]], [com ? `Diferenças de texto em ${com} página(s).` : 'Os textos são iguais. Obs.: só o texto é comparado, não a aparência.']];
} };
F.img2pdf = { min: 1, max: 100, ext: IMG, async fn(arqs, op) {
  const tam = op.tamanho || 'a4', orient = op.orientacao || 'auto', marg = num(op.margem, tam !== 'imagem' ? 10 : 0, 0, 100) * MM, sep = !!op.separado;
  const unico = await LibDoc.create(), s = [];
  for (const [nome, dados] of arqs) {
    const bm = await abrirImg(nome, dados), jpg = await canvasPara(desenhar(bm, bm.width, bm.height), 'jpg', .92);
    const alvo = sep ? await LibDoc.create() : unico, im = await alvo.embedJpg(jpg);
    let pw, ph; if (tam === 'imagem') { pw = bm.width * 72 / 96; ph = bm.height * 72 / 96; } else { [pw, ph] = tam === 'carta' ? [612, 792] : [595.28, 841.89]; if (orient === 'paisagem' || (orient === 'auto' && bm.width > bm.height)) [pw, ph] = [ph, pw]; }
    const pg = alvo.addPage([pw, ph]), aw = pw - 2 * marg, ah = ph - 2 * marg, esc = Math.min(aw / bm.width, ah / bm.height), w = bm.width * esc, h = bm.height * esc;
    pg.drawImage(im, { x: marg + (aw - w) / 2, y: marg + (ah - h) / 2, width: w, height: h });
    if (sep) s.push([base(nome) + '.pdf', await alvo.save()]);
  }
  return [sep ? s : [['imagens.pdf', await unico.save()]], []];
} };
F.pdf2img = { min: 1, max: 10, ext: PDF, fn(arqs, op) {
  const fmt = op.formato === 'png' ? 'png' : 'jpg', dpi = Math.round(num(op.dpi, 150, 50, 300)), s = [];
  for (const [nome, dados] of arqs) {
    const d = abrir(dados, nome);
    for (const i of faixa(op.paginas, d.countPages(), true)) { const pix = d.loadPage(i).toPixmap(M.Matrix.scale(dpi / 72, dpi / 72), M.ColorSpace.DeviceRGB, false);
      s.push([`${base(nome)}-pagina-${String(i + 1).padStart(3, '0')}.${fmt}`, (fmt === 'jpg' ? pix.asJPEG(90) : pix.asPNG()).slice()]); pix.destroy(); }
    d.destroy();
  }
  return [s, [`${s.length} imagem(ns) geradas a ${dpi} dpi.`]];
} };
F.pdf2txt = { min: 1, max: 20, ext: PDF, fn(arqs, op) {
  const s = arqs.map(([nome, dados]) => { const d = abrir(dados, nome); const t = faixa(op.paginas, d.countPages(), true).map(i => { const st = d.loadPage(i).toStructuredText(''); const x = st.asText(); st.destroy(); return x; }).join('\n\n'); d.destroy(); return [base(nome) + '.txt', new TextEncoder().encode(t)]; });
  return [s, s.some(x => x[1].length > 4) ? [] : ['Nenhum texto encontrado: o PDF parece escaneado. Use o OCR antes.']];
} };
F.tarja_auto = { min: 1, max: 20, ext: PDF, async fn(arqs, op) {
  const regs = expressoes(op), s = [], a = [];
  for (const [nome, dados] of arqs) {
    const d = abrir(dados, nome), cont = {}, alvo = faixa(op.paginas, d.countPages(), true), caixas = []; let semTexto = 0;
    for (const i of alvo) {
      const pg = d.loadPage(i), inv = M.Matrix.invert(pg.getTransform());
      const st = pg.toStructuredText(''); if (!st.asText().trim()) semTexto++; st.destroy();
      const ach = ocorrencias(pg, regs);
      for (const [rot, , r] of ach) { const ann = pg.createAnnotation('Redact'); ann.setRect([r[0] - .5, r[1] - .5, r[2] + .5, r[3] + .5]); cont[rot] = (cont[rot] || 0) + 1; caixas.push([i, rectPdf(inv, [r[0] - .5, r[1] - .5, r[2] + .5, r[3] + .5])]); }
      if (ach.length) pg.applyRedactions(false, M.PDFPage.REDACT_IMAGE_PIXELS, M.PDFPage.REDACT_LINE_ART_NONE, M.PDFPage.REDACT_TEXT_REMOVE);
    }
    if (op.metadados !== false) { for (const k of ['info:Title', 'info:Author', 'info:Subject', 'info:Keywords', 'info:Creator', 'info:Producer']) try { d.setMetaData(k, ''); } catch (e) {}
      try { d.getTrailer().get('Root').delete('Metadata'); } catch (e) {} }
    const pdfLimpo = salvar(d, 'garbage=compact,compress,clean'); d.destroy();
    const pd = await LibDoc.load(pdfLimpo), pages = pd.getPages(), cor = corPdf(op.cor, [0, 0, 0]);
    for (const [i, r] of caixas) pages[i].drawRectangle({ x: r.x, y: r.y, width: r.w, height: r.h, color: cor });
    s.push([base(nome) + '-tarjado.pdf', await pd.save()]);
    const total = Object.values(cont).reduce((x, y) => x + y, 0);
    a.push(total ? `“${nome}”: ${total} trecho(s) apagado(s) de verdade (${Object.entries(cont).map(([k, v]) => k + ': ' + v).join(', ')}).` : `“${nome}”: nada encontrado.`);
    if (semTexto) a.push(`“${nome}”: ${semTexto} página(s) sem texto (escaneadas). Use o OCR antes, senão os dados nelas não são achados.`);
  }
  return [s, a];
} };
F.substituir = { min: 1, max: 20, ext: PDF, async fn(arqs, op) {
  const busca = String(op.busca || ''); if (!busca.trim()) throw new Erro('Digite o texto a procurar.');
  const troca = String(op.troca || ''), rx = new RegExp((op.inteira ? String.raw`(?<![\p{L}\p{N}_])` : '') + escapar(busca) + (op.inteira ? String.raw`(?![\p{L}\p{N}_])` : ''), 'gu' + (op.maiusculas ? '' : 'i'));
  const s = [], a = [];
  for (const [nome, dados] of arqs) {
    const d = abrir(dados, nome), ops = []; let trocas = 0, giradas = 0, mistas = 0;
    for (const pno of faixa(op.paginas, d.countPages(), true)) {
      const pg = d.loadPage(pno), rot = rotacao(pg);
      for (const b of blocosDaPagina(pg)) for (const l of b.lines) {
        const txt = l.chars.map(c => c.c).join(''), n = (txt.match(rx) || []).length; if (!n) continue;
        if (rot || Math.abs(l.dir[1]) > .02 || l.dir[0] < .98) { giradas += n; continue; }
        const vis = l.chars.filter(c => c.c.trim()), main = vis[0] || l.chars[0];
        if (new Set(vis.map(c => semSubset(c.f.name) + Math.round(c.size) + c.color)).size > 1) mistas++;
        ops.push({ kind: 'rep', page: pno, bbox: l.bbox, origin: [vis[0] ? vis[0].rect[0] : l.bbox[0], main.origin[1]], size: main.size, color: main.color,
          family: familia(main.f), bold: ehNegrito(main.f), italic: ehItalico(main.f), text: txt.replace(rx, () => troca).trim(), align: 'left' });
        trocas += n;
      }
    }
    d.destroy();
    let saida = dados, av = [];
    if (ops.length) { const r = await aplicar(dados, ops); saida = b64ParaBytes(r.pdf); av = r.warnings; }
    s.push([base(nome) + '-substituido.pdf', saida]);
    a.push(trocas ? `“${nome}”: ${trocas} ocorrência(s) trocada(s).` : `“${nome}”: “${busca}” não foi encontrado.`);
    if (giradas) a.push(`${giradas} ocorrência(s) em texto girado ou inclinado não foram trocadas.`);
    if (mistas) a.push(`${mistas} linha(s) misturavam estilos (negrito, cores ou tamanhos) e ficaram com um estilo só.`);
    a.push(...av);
  }
  return [s, a];
} };
F.img_comprimir = { min: 1, max: 100, ext: IMG, async fn(arqs, op) {
  const q = { baixa: .85, media: .72, alta: .55 }[op.nivel] || .72, ml = Math.round(num(op.max_lado, 0, 0, 20000)), s = [], a = [];
  for (const [nome, dados] of arqs) {
    const bm = await abrirImg(nome, dados), esc = ml && Math.max(bm.width, bm.height) > ml ? ml / Math.max(bm.width, bm.height) : 1;
    let fmt = fmtDe(nome); if (fmt === 'png' && op.nivel === 'alta') fmt = 'webp';
    let novo = await canvasPara(desenhar(bm, Math.round(bm.width * esc), Math.round(bm.height * esc)), fmt, q), ext = fmt;
    if (novo.length >= dados.length && esc === 1) { novo = dados; ext = (nome.split('.').pop() || 'png').toLowerCase(); a.push(`“${nome}” já estava otimizada: mantive a original.`); }
    else a.push(`“${nome}”: de ${kb(dados.length)} para ${kb(novo.length)} (${Math.max(0, Math.round(100 - 100 * novo.length / dados.length))}% menor).`);
    s.push([`${base(nome)}-comprimida.${ext}`, novo]);
  }
  return [s, a];
} };
F.img_redimensionar = { min: 1, max: 100, ext: IMG, async fn(arqs, op) {
  const s = [];
  for (const [nome, dados] of arqs) {
    const bm = await abrirImg(nome, dados), W = bm.width, H = bm.height; let w, h;
    if ((op.modo || 'porcentagem') === 'porcentagem') { const p = num(op.porcentagem, 50, 1, 1000) / 100; w = W * p; h = H * p; }
    else { const lw = Math.round(num(op.largura, 0, 0, 20000)), lh = Math.round(num(op.altura, 0, 0, 20000)); if (!lw && !lh) throw new Erro('Informe a largura ou a altura em pixels.');
      if (op.proporcao !== false) { const e = Math.min(lw ? lw / W : 1e9, lh ? lh / H : 1e9); w = W * e; h = H * e; } else { w = lw || W; h = lh || H; } }
    if (op.nao_aumentar !== false && (w > W || h > H)) { const e = Math.min(W / w, H / h); w *= e; h *= e; }
    w = Math.max(1, Math.round(w)); h = Math.max(1, Math.round(h));
    s.push([`${base(nome)}-${w}x${h}.${fmtDe(nome)}`, await canvasPara(desenhar(bm, w, h), fmtDe(nome), .9)]);
  }
  return [s, []];
} };
F.img_recortar = { min: 1, max: 100, ext: IMG, async fn(arqs, op) {
  const s = [];
  for (const [nome, dados] of arqs) {
    const bm = await abrirImg(nome, dados), W = bm.width, H = bm.height; let c;
    if (op.modo === 'margens') { const [t, b, e, d] = ['topo', 'base', 'esq', 'dir'].map(k => Math.round(num(op[k], 0, 0, 20000))); c = [e, t, W - d, H - b]; }
    else { const [a, b] = { '1:1': [1, 1], '4:3': [4, 3], '3:4': [3, 4], '16:9': [16, 9], '9:16': [9, 16], '3:2': [3, 2], '2:3': [2, 3], a4: [210, 297] }[op.proporcao || '1:1'] || [1, 1];
      if (W / H > a / b) { const nw = Math.round(H * a / b), x = (W - nw) >> 1; c = [x, 0, x + nw, H]; } else { const nh = Math.round(W * b / a), y = (H - nh) >> 1; c = [0, y, W, y + nh]; } }
    if (c[2] - c[0] < 2 || c[3] - c[1] < 2) throw new Erro(`O recorte de “${nome}” ficou vazio: as margens são maiores que a imagem (${W}x${H}).`);
    s.push([`${base(nome)}-recortada.${fmtDe(nome)}`, await canvasPara(desenhar(bm, c[2] - c[0], c[3] - c[1], c[0], c[1], c[2] - c[0], c[3] - c[1]), fmtDe(nome), .92)]);
  }
  return [s, []];
} };
F.img_converter = { min: 1, max: 100, ext: IMG, async fn(arqs, op) {
  const fmt = op.formato || 'jpg'; if (!['jpg', 'png', 'webp'].includes(fmt)) throw new Erro('No celular dá para converter para JPG, PNG ou WEBP.');
  const q = num(op.qualidade, 90, 10, 100) / 100, s = [];
  for (const [nome, dados] of arqs) { const bm = await abrirImg(nome, dados); s.push([`${base(nome)}.${fmt}`, await canvasPara(desenhar(bm, bm.width, bm.height), fmt, q)]); }
  return [s, []];
} };
F.img_girar = { min: 1, max: 100, ext: IMG, async fn(arqs, op) {
  const ang = Math.round(num(op.angulo, 90, 0, 270)), s = [];
  for (const [nome, dados] of arqs) { const bm = await abrirImg(nome, dados); s.push([`${base(nome)}-girada.${fmtDe(nome)}`, await canvasPara(desenhar(bm, bm.width, bm.height, 0, 0, bm.width, bm.height, ang, op.espelhar || ''), fmtDe(nome), .92)]); }
  return [s, []];
} };

const MIME = { pdf: 'application/pdf', zip: 'application/zip', jpg: 'image/jpeg', png: 'image/png', webp: 'image/webp', txt: 'text/plain; charset=utf-8' };
async function executar(ferramenta, arquivos, opcoes) {
  const t = F[ferramenta];
  if (!t) throw new Erro('Esta ferramenta não está disponível no celular.');
  if (!(arquivos.length >= t.min && arquivos.length <= t.max)) throw new Erro(t.min === t.max ? `Esta ferramenta precisa de ${t.min} arquivo(s).` : `Esta ferramenta aceita de ${t.min} a ${t.max} arquivo(s).`);
  const arqs = arquivos.map(f => {
    const nome = String(f.nome || 'arquivo'), dados = b64ParaBytes(f.b64);
    if (!t.ext.some(e => nome.toLowerCase().endsWith(e))) throw new Erro(`“${nome}”: tipo de arquivo não aceito (${t.ext.join(', ')}).`);
    if (!dados.length) throw new Erro(`“${nome}” está vazio.`);
    return [nome, dados];
  });
  let [saidas, avisos] = await t.fn(arqs, opcoes || {});
  if (!saidas.length) throw new Erro('Nada foi gerado.');
  if (saidas.length > 1) {
    const usados = new Set();
    saidas = [[`${ferramenta}-resultado.zip`, zip(saidas.map(([n, b]) => { let u = n, k = 1; while (usados.has(u)) { k++; u = n.replace(/(\.[^.]*)?$/, ` (${k})$1`); } usados.add(u); return [u, b]; }))]];
  }
  const [nome, bytes] = saidas[0];
  return { nome, mime: MIME[(nome.split('.').pop() || '').toLowerCase()] || 'application/octet-stream', b64: bytesParaB64(bytes), avisos };
}

/* ------------------------------------------------------------ rotas (mesmas do servidor.py) */
export async function chamar(caminho, corpo = {}) {
  try {
    if (caminho === '/api/ping') return { ok: true, motor: 'celular', pymupdf: 'MuPDF.js', rede: false, docx: false, ocr: true, idiomas: ['eng', 'por'], office: [], navegador: false, pdfa: false, assinatura: false };
    if (caminho === '/api/tool') return await executar(String(corpo.ferramenta), corpo.arquivos || [], corpo.opcoes || {});
    const pdf = b64ParaBytes(corpo.pdf || '');
    if (caminho === '/api/analyze') return analisar(pdf, +corpo.page);
    if (caminho === '/api/apply') return await aplicar(pdf, corpo.ops);
    if (caminho === '/api/ocr') return await ocr(pdf, corpo.paginas, corpo.idioma, corpo.pular_com_texto !== false);
    if (caminho === '/api/buscar') return { resultados: buscar(pdf, corpo.opcoes || {}) };
    throw new Erro('Esta função não está disponível no celular.');
  } catch (e) {
    if (!(e instanceof Erro)) console.error(e);
    throw new Error(e instanceof Erro ? e.message : 'Erro no motor de PDF: ' + (e && e.message || e));
  }
}
