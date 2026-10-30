/* Funções compartilhadas pela matriz e pelo app do técnico */
'use strict';
const CFG = window.OSC_CONFIG || {};
const API_OK = !!CFG.api && /^https?:\/\//.test(CFG.api);

function uid() {
  const b = new Uint8Array(12); crypto.getRandomValues(b);
  return Date.now().toString(36) + Array.from(b, x => x.toString(36).padStart(2, '0')).join('').slice(0, 14);
}
const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const clone = o => JSON.parse(JSON.stringify(o));
const hoje = () => { const d = new Date(); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`; };
const dataBR = s => { const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})/); return m ? `${m[3]}/${m[2]}/${m[1]}` : ''; };
const dataHoraBR = s => { if (!s) return ''; const d = new Date(s); return isNaN(d) ? '' : d.toLocaleString('pt-BR', { day: '2-digit', month: '2-digit', hour: '2-digit', minute: '2-digit' }); };
function lsGet(k, def) { try { const v = localStorage.getItem(k); return v == null ? def : JSON.parse(v); } catch (e) { return def; } }
function lsSet(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); return true; } catch (e) { return false; } }
function debounce(fn, ms) { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); }; }

/* ---------- conversa com o Apps Script ---------- */
class ErroRede extends Error { }
let CHAVE = '';
async function pedido(url, opcoes) {
  let r;
  try { r = await fetch(url, Object.assign({ cache: 'no-store', redirect: 'follow' }, opcoes)); }
  catch (e) { throw new ErroRede('sem conexão'); }
  let j;
  try { j = await r.json(); }
  catch (e) { throw new ErroRede('O servidor não respondeu direito (sinal fraco, ou implantação do Apps Script sem acesso "Qualquer pessoa").'); }
  if (j && j.erro) { const e = new Error(j.erro); e.status = j.status || 500; throw e; }
  return j;
}
function apiGet(acao, params) {
  const q = new URLSearchParams(Object.assign({ a: acao, k: CHAVE }, params || {}));
  return pedido(CFG.api + '?' + q.toString(), { method: 'GET' });
}
function apiPost(acao, dados) {
  // text/plain evita a checagem prévia (CORS) que o Apps Script não atende
  return pedido(CFG.api, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(Object.assign({ a: acao, k: CHAVE }, dados || {})) });
}

/* ---------- fotos ---------- */
function blobParaB64(blob) { return new Promise((ok, falha) => { const r = new FileReader(); r.onload = () => ok(String(r.result).split(',')[1] || ''); r.onerror = () => falha(r.error); r.readAsDataURL(blob); }); }
function b64ParaBlob(b64, tipo) { const s = atob(b64); const u = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) u[i] = s.charCodeAt(i); return new Blob([u], { type: tipo || 'image/jpeg' }); }
const FOTOS = new Map();             // id -> objectURL
const FOTOS_BLOB = new Map();        // id -> Blob
const _baixando = new Map();
async function fotoBlob(id, guardar) {
  if (FOTOS_BLOB.has(id)) return FOTOS_BLOB.get(id);
  if (_baixando.has(id)) return _baixando.get(id);
  const p = apiGet('foto', { id }).then(r => { const b = b64ParaBlob(r.b64); registrarFoto(id, b); if (guardar) guardar(id, b); return b; }).finally(() => _baixando.delete(id));
  _baixando.set(id, p); return p;
}
function registrarFoto(id, blob) { if (!FOTOS.has(id)) { FOTOS_BLOB.set(id, blob); FOTOS.set(id, URL.createObjectURL(blob)); } return FOTOS.get(id); }
/* Preenche <img data-foto="id"> ainda sem imagem, até 4 downloads de cada vez */
let _filaHidr = [], _ativos = 0;
function hidratarFotos(raiz, guardar) {
  raiz.querySelectorAll('img[data-foto]').forEach(img => { const id = img.dataset.foto; if (FOTOS.has(id)) { if (img.src !== FOTOS.get(id)) img.src = FOTOS.get(id); } else if (!_filaHidr.includes(id)) _filaHidr.push(id); });
  const rodar = () => {
    while (_ativos < 4 && _filaHidr.length) {
      const id = _filaHidr.shift(); _ativos++;
      fotoBlob(id, guardar).then(() => { document.querySelectorAll(`img[data-foto="${CSS.escape(id)}"]`).forEach(i => { i.src = FOTOS.get(id); }); }).catch(() => { }).finally(() => { _ativos--; rodar(); });
    }
  };
  rodar();
}

/* Re-renderiza mantendo foco, cursor e textos em digitação */
function pintar(el, html) {
  const a = document.activeElement; let chave = null, s0, s1, valor = null;
  if (a && el.contains(a) && a.dataset && a.dataset.f) { chave = a.dataset.f; try { s0 = a.selectionStart; s1 = a.selectionEnd; } catch (e) { } if ((a.tagName === 'TEXTAREA' || a.tagName === 'INPUT') && a.type !== 'checkbox') valor = a.value; }
  const rasc = {}; el.querySelectorAll('[data-rasc]').forEach(n => { if (n.value) rasc[n.dataset.f] = n.value; });
  el.innerHTML = html;
  el.querySelectorAll('[data-rasc]').forEach(n => { if (rasc[n.dataset.f] && !n.value) n.value = rasc[n.dataset.f]; });
  if (chave) { const n = el.querySelector(`[data-f="${CSS.escape(chave)}"]`); if (n) { if (valor != null) n.value = valor; n.focus({ preventScroll: true }); try { if (s0 != null) n.setSelectionRange(s0, s1); } catch (e) { } } }
}

let _toastT;
function toast(msg, ms = 2600) {
  const t = document.getElementById('toast'); if (!t) return;
  t.textContent = msg; t.hidden = false; clearTimeout(_toastT); _toastT = setTimeout(() => { t.hidden = true; }, ms);
}
function perguntar(titulo, texto, botoes) {
  return new Promise(ok => {
    const o = document.createElement('div'); o.className = 'overlay';
    o.innerHTML = `<div class="sheet" role="dialog" aria-modal="true"><h2>${esc(titulo)}</h2>${texto ? `<div class="muted">${texto}</div>` : ''}<div class="acoes">${botoes.map((b, i) => `<button class="btn ${b.cls || ''}" data-i="${i}">${esc(b.txt)}</button>`).join('')}</div></div>`;
    o.addEventListener('click', e => { const b = e.target.closest('[data-i]'); if (b) { o.remove(); ok(botoes[+b.dataset.i].v); } else if (e.target === o) { o.remove(); ok(null); } });
    document.body.appendChild(o);
    const ult = o.querySelectorAll('button'); if (ult.length) ult[ult.length - 1].focus();
  });
}
function contarOS(os, respDe) {
  let tot = 0, ok = 0, nok = 0;
  for (const l of os.locais || []) for (const e of l.etapas || []) for (const it of e.itens || []) { tot++; const r = respDe(it.id); if (r && r.status === 'ok') ok++; else if (r && r.status === 'nok') nok++; }
  return { tot, ok, nok, feitos: ok + nok };
}
const TELA_SEM_CONFIG = `<div class="card empty" style="margin-top:20px"><b>Falta configurar o endereço do Apps Script.</b><br>Abra o arquivo <span class="mono">config.js</span>, cole a URL que termina em <span class="mono">/exec</span> e publique de novo.</div>`;
