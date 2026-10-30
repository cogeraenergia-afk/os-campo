/* ============================================================================
   Relatório Word (.docx) da OS — gerado no navegador, sem bibliotecas externas.
   Estrutura: capa Cogera (folder) · conteúdo sobre o timbrado · fechamento.
   Tudo editável no Word: textos, tabelas e fotos.
   ============================================================================ */
'use strict';
(function () {

/* ---------- ZIP (sem compressão; o Word abre normalmente) ---------- */
const CRC_T = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
function crc32(b) { let c = 0xFFFFFFFF; for (let i = 0; i < b.length; i++) c = CRC_T[(c ^ b[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
const _te = new TextEncoder();
function zip(entradas) {
  const partes = [], central = []; let off = 0;
  const d = new Date(); const dt = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate(); const tm = (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1);
  for (const e of entradas) {
    const nome = _te.encode(e.nome);
    const dados = e.dados instanceof Uint8Array ? e.dados : _te.encode(e.dados);
    const crc = crc32(dados);
    const h = new DataView(new ArrayBuffer(30)); h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true); h.setUint16(8, 0, true); h.setUint16(10, tm, true); h.setUint16(12, dt, true); h.setUint32(14, crc, true); h.setUint32(18, dados.length, true); h.setUint32(22, dados.length, true); h.setUint16(26, nome.length, true); h.setUint16(28, 0, true);
    partes.push(h.buffer, nome, dados);
    const c = new DataView(new ArrayBuffer(46)); c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x0800, true); c.setUint16(10, 0, true); c.setUint16(12, tm, true); c.setUint16(14, dt, true); c.setUint32(16, crc, true); c.setUint32(20, dados.length, true); c.setUint32(24, dados.length, true); c.setUint16(28, nome.length, true); c.setUint32(42, off, true);
    central.push(c.buffer, nome);
    off += 30 + nome.length + dados.length;
  }
  let tamCD = 0; central.forEach(p => { tamCD += p.byteLength; });
  const f = new DataView(new ArrayBuffer(22)); f.setUint32(0, 0x06054b50, true); f.setUint16(8, entradas.length, true); f.setUint16(10, entradas.length, true); f.setUint32(12, tamCD, true); f.setUint32(16, off, true);
  return new Blob([...partes, ...central, f.buffer], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
}

/* ---------- JPEG: largura e altura ---------- */
function tamanhoJPEG(b) {
  let i = 2;
  while (i < b.length) {
    if (b[i] !== 0xFF) { i++; continue; }
    const m = b[i + 1];
    if (m >= 0xC0 && m <= 0xCF && m !== 0xC4 && m !== 0xC8 && m !== 0xCC) return { h: (b[i + 5] << 8) | b[i + 6], w: (b[i + 7] << 8) | b[i + 8] };
    if (m === 0xD8 || m === 0x01 || (m >= 0xD0 && m <= 0xD7)) { i += 2; continue; }
    i += 2 + ((b[i + 2] << 8) | b[i + 3]);
  }
  return { w: 4, h: 3 };
}

/* ---------- XML ---------- */
const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
const CM = 360000;            // EMU por cm
const TW = 567;               // twips por cm
const NAVY = '26306E', GREEN = '1E9A4B';
const NS = 'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';

function run(texto, o = {}) {
  const rp = [o.b ? '<w:b/>' : '', o.i ? '<w:i/>' : '', o.cor ? `<w:color w:val="${o.cor}"/>` : '', o.sz ? `<w:sz w:val="${o.sz * 2}"/><w:szCs w:val="${o.sz * 2}"/>` : '', o.hl ? `<w:highlight w:val="${o.hl}"/>` : '', o.caps ? '<w:caps/>' : ''].join('');
  const linhas = String(texto == null ? '' : texto).split(/\r?\n/);
  return `<w:r>${rp ? `<w:rPr>${rp}</w:rPr>` : ''}${linhas.map((l, i) => (i ? '<w:br/>' : '') + `<w:t xml:space="preserve">${esc(l)}</w:t>`).join('')}</w:r>`;
}
function par(conteudo, o = {}) {
  const pp = [o.estilo ? `<w:pStyle w:val="${o.estilo}"/>` : '', o.keep ? '<w:keepNext/>' : '', o.quebra ? '<w:pageBreakBefore/>' : '', o.esp != null ? `<w:spacing w:before="${o.antes || 0}" w:after="${o.esp}"/>` : '', o.al ? `<w:jc w:val="${o.al}"/>` : ''].join('');
  const c = typeof conteudo === 'string' && !conteudo.startsWith('<w:') ? run(conteudo, o) : (conteudo || '');
  return `<w:p>${pp ? `<w:pPr>${pp}</w:pPr>` : ''}${c}</w:p>`;
}
function celula(conteudo, largura, o = {}) {
  const tc = [`<w:tcW w:w="${largura}" w:type="dxa"/>`, o.fill ? `<w:shd w:val="clear" w:color="auto" w:fill="${o.fill}"/>` : '', o.span ? `<w:gridSpan w:val="${o.span}"/>` : '', o.semBorda ? '<w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders>' : '', '<w:vAlign w:val="' + (o.va || 'top') + '"/>'].join('');
  const ps = Array.isArray(conteudo) ? conteudo.join('') : par(conteudo, Object.assign({ esp: 0 }, o.p || {}));
  return `<w:tc><w:tcPr>${tc}</w:tcPr>${ps}</w:tc>`;
}
function tabela(linhas, larguras, o = {}) {
  const total = larguras.reduce((a, b) => a + b, 0);
  const bordas = o.semBorda ? '<w:tblBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders>'
    : '<w:tblBorders><w:top w:val="single" w:sz="4" w:color="C3CADB"/><w:left w:val="single" w:sz="4" w:color="C3CADB"/><w:bottom w:val="single" w:sz="4" w:color="C3CADB"/><w:right w:val="single" w:sz="4" w:color="C3CADB"/><w:insideH w:val="single" w:sz="4" w:color="C3CADB"/><w:insideV w:val="single" w:sz="4" w:color="C3CADB"/></w:tblBorders>';
  const tr = linhas.map((l, i) => `<w:tr><w:trPr><w:cantSplit/>${i === 0 && o.cabecalho ? '<w:tblHeader/>' : ''}</w:trPr>${l}</w:tr>`).join('');
  return `<w:tbl><w:tblPr><w:tblW w:w="${total}" w:type="dxa"/><w:tblLayout w:type="fixed"/>${bordas}<w:tblCellMar><w:top w:w="50" w:type="dxa"/><w:left w:w="90" w:type="dxa"/><w:bottom w:w="50" w:type="dxa"/><w:right w:w="90" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>${larguras.map(w => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>${tr}</w:tbl>`;
}
const cab = (t, w) => celula(t, w, { fill: NAVY, p: { b: true, cor: 'FFFFFF', sz: 9 } });

let docPrId = 1;
function grafico(rid, cx, cy, nome) {
  const id = docPrId++;
  return `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="${esc(nome)}"/><pic:cNvPicPr><a:picLocks noChangeAspect="1"/></pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>`;
}
function imgInline(rid, cx, cy, nome) {
  const id = docPrId;
  return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${id}" name="${esc(nome)}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>${grafico(rid, cx, cy, nome)}</wp:inline></w:drawing></w:r>`;
}
function imgFundo(rid, nome) {
  const id = docPrId; const cx = Math.round(21 * CM), cy = Math.round(29.7 * CM);
  return `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${id}" behindDoc="1" locked="1" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="${id}" name="${esc(nome)}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>${grafico(rid, cx, cy, nome)}</wp:anchor></w:drawing></w:r>`;
}

/* ---------- textos ---------- */
const SIT = { ok: ['Conforme', 'E2F3E8', '17803F'], nok: ['Não conforme', 'FBE3E1', 'B8261F'], '': ['Não verificado', 'ECEEF4', '6D7490'] };
function dataBR(s) { if (!s) return ''; const m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})/); return m ? `${m[3]}/${m[2]}/${m[1]}` : String(s); }
function dataHoraBR(s) { if (!s) return ''; const d = new Date(s); if (isNaN(d)) return String(s); return d.toLocaleString('pt-BR', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }); }

/* ============================================================================ */
async function gerarRelatorio({ os, respostas, fotos, tecnicos, lerFoto, lerModelo, comCapa, empresa }) {
  docPrId = 1;
  const midia = [];            // {nome, dados}
  const relsDoc = [];          // relações do documento
  let ridN = 10;
  const addImg = (dados, ext = 'jpg') => { const nome = `img${midia.length + 1}.${ext}`; midia.push({ nome, dados }); const rid = 'rId' + (ridN++); relsDoc.push(`<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${nome}"/>`); return { rid, nome }; };

  const resp = new Map(respostas.map(r => [r.itemId, r]));
  const fotosItem = new Map();
  for (const f of fotos.slice().sort((a, b) => String(a.em).localeCompare(String(b.em)))) { if (!fotosItem.has(f.itemId)) fotosItem.set(f.itemId, []); fotosItem.get(f.itemId).push(f); }
  const nomesTec = (os.tecnicos || []).map(id => (tecnicos.find(t => t.id === id) || {}).nome).filter(Boolean).join(', ');
  const locais = (os.locais || []).filter(l => l);
  const LARG = Math.round(17.1 * TW); // largura útil sobre o timbrado

  const MODELO = {};
  for (const n of ['timbrado.jpg', ...(comCapa ? ['capa-1.jpg', 'capa-2.jpg', 'capa-3.jpg', 'capa-4.jpg', 'fechamento-1.jpg', 'fechamento-2.jpg'] : [])]) { try { MODELO[n] = await lerModelo(n); } catch (e) { MODELO[n] = null; } }
  const ler = f => MODELO[f] || null;
  const corpo = [];

  /* capa (folder Cogera) */
  const capas = comCapa ? ['capa-1.jpg', 'capa-2.jpg', 'capa-3.jpg', 'capa-4.jpg'].map(ler).filter(Boolean) : [];
  capas.forEach((b, i) => { const { rid, nome } = addImg(b); corpo.push(`<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>${imgFundo(rid, nome)}${i < capas.length - 1 ? '<w:r><w:br w:type="page"/></w:r>' : ''}</w:p>`); });
  const secCapa = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="0" w:right="0" w:bottom="0" w:left="0" w:header="0" w:footer="0" w:gutter="0"/>';
  if (capas.length) corpo.push(`<w:p><w:pPr><w:spacing w:before="0" w:after="0"/><w:sectPr><w:headerReference w:type="default" r:id="rId4"/><w:footerReference w:type="default" r:id="rId6"/>${secCapa}</w:sectPr></w:pPr></w:p>`);

  /* identificação */
  corpo.push(par(run('Relatório de serviço', { b: true, sz: 20, cor: NAVY }), { esp: 60 }));
  corpo.push(par(run(os.titulo || 'Manutenção preventiva', { sz: 13, cor: GREEN, b: true }), { esp: 240 }));
  const ident = [
    ['Ordem de serviço', os.numero || ''], ['Serviço', os.titulo || ''], ['Cliente', os.cliente || ''],
    ['Data de execução', dataBR(os.data)], ['Técnico(s)', nomesTec], ['Executante', empresa || ''],
    ['Situação da OS', os.status === 'concluida' ? `Concluída em ${dataHoraBR(os.concluidaEm)}` : 'Em andamento'],
  ];
  corpo.push(tabela(ident.map(([k, v]) => celula(k, Math.round(LARG * 0.3), { fill: 'F5F7FB', p: { b: true, sz: 9.5, cor: NAVY } }) + celula(v, LARG - Math.round(LARG * 0.3), { p: { sz: 9.5 } })), [Math.round(LARG * 0.3), LARG - Math.round(LARG * 0.3)]));

  /* 1. resumo */
  let tot = 0, ok = 0, nok = 0, pend = 0; const ncs = []; const porLocal = [];
  for (const l of locais) {
    let lt = 0, lo = 0, ln = 0;
    for (const e of l.etapas || []) for (const it of e.itens || []) {
      const r = resp.get(it.id) || {}; lt++; tot++;
      if (r.status === 'ok') { lo++; ok++; } else if (r.status === 'nok') { ln++; nok++; ncs.push({ l, e, it, r }); } else pend++;
    }
    porLocal.push([l.nome, lt, lo, ln, lt - lo - ln]);
  }
  corpo.push(par('1. Resumo', { estilo: 'Heading1', quebra: false }));
  const frase = `Foram verificados ${tot} ${tot === 1 ? 'item' : 'itens'} em ${locais.length} ${locais.length === 1 ? 'local' : 'locais'}: ${ok} conforme${ok === 1 ? '' : 's'}, ${nok} não conforme${nok === 1 ? '' : 's'}${pend ? ` e ${pend} não verificado${pend === 1 ? '' : 's'}` : ''}.`;
  corpo.push(par(frase, { esp: 120 }));
  const wR = [LARG - 4 * 1500, 1500, 1500, 1500, 1500];
  corpo.push(tabela([
    cab('Local', wR[0]) + cab('Itens', wR[1]) + cab('Conformes', wR[2]) + cab('Não conformes', wR[3]) + cab('Não verificados', wR[4]),
    ...porLocal.map(([n, a, b, c, d]) => celula(n, wR[0], { p: { sz: 9.5 } }) + celula(String(a), wR[1], { p: { sz: 9.5, al: 'center' } }) + celula(String(b), wR[2], { p: { sz: 9.5, al: 'center', cor: b ? '17803F' : undefined } }) + celula(String(c), wR[3], { p: { sz: 9.5, al: 'center', b: !!c, cor: c ? 'B8261F' : undefined } }) + celula(String(d), wR[4], { p: { sz: 9.5, al: 'center' } })),
  ], wR, { cabecalho: true }));

  /* 2. não conformidades */
  corpo.push(par('2. Não conformidades', { estilo: 'Heading1' }));
  if (!ncs.length) corpo.push(par('Nenhuma não conformidade registrada nesta OS.'));
  else {
    const wN = [700, 3300, 2400, LARG - 6400];
    corpo.push(tabela([
      cab('Nº', wN[0]) + cab('Local / etapa', wN[1]) + cab('Item', wN[2]) + cab('Observação do técnico', wN[3]),
      ...ncs.map((x, i) => celula(String(i + 1), wN[0], { p: { sz: 9.5, al: 'center' } }) + celula(`${x.l.nome} · ${x.e.nome}`, wN[1], { p: { sz: 9.5 } }) + celula(x.it.nome, wN[2], { p: { sz: 9.5, b: true } }) + celula(x.r.obs || '—', wN[3], { p: { sz: 9.5 } })),
    ], wN, { cabecalho: true }));
  }

  /* 3. registros por local e etapa */
  corpo.push(par('3. Registros por local', { estilo: 'Heading1' }));
  for (let li = 0; li < locais.length; li++) {
    const l = locais[li];
    corpo.push(par(`3.${li + 1} ${l.nome}`, { estilo: 'Heading2' }));
    const etapas = l.etapas || [];
    if (!etapas.length) corpo.push(par('Nenhuma etapa cadastrada.', { i: true }));
    for (let ei = 0; ei < etapas.length; ei++) {
      const e = etapas[ei]; const itens = e.itens || [];
      corpo.push(par(`3.${li + 1}.${ei + 1} ${e.nome}`, { estilo: 'Heading3' }));
      if (!itens.length) { corpo.push(par('Nenhum item cadastrado.', { i: true })); continue; }
      const temValor = itens.some(it => it.unid || (resp.get(it.id) || {}).valor);
      const w = temValor ? [3200, 1700, 1500, LARG - 6400] : [3600, 1800, LARG - 5400];
      const linhas = [cab('Item', w[0]) + cab('Situação', w[1]) + (temValor ? cab('Valor', w[2]) : '') + cab('Observação', w[w.length - 1])];
      for (const it of itens) {
        const r = resp.get(it.id) || {}; const [txt, fill, cor] = SIT[r.status || ''] || SIT[''];
        const val = r.valor != null && r.valor !== '' ? `${String(r.valor).replace('.', ',')}${it.unid ? ' ' + it.unid : ''}` : '—';
        linhas.push(celula(it.nome, w[0], { p: { sz: 9.5 } }) + celula(txt, w[1], { fill, p: { sz: 9.5, b: true, cor } }) + (temValor ? celula(val, w[2], { p: { sz: 9.5 } }) : '') + celula(r.obs || '', w[w.length - 1], { p: { sz: 9.5 } }));
      }
      corpo.push(tabela(linhas, w, { cabecalho: true }));

      /* registro fotográfico da etapa: 2 fotos por linha */
      const fs_ = [];
      for (const it of itens) { const lst = fotosItem.get(it.id) || []; lst.forEach((f, k) => fs_.push({ f, legenda: `${it.nome}${lst.length > 1 ? ` — foto ${k + 1} de ${lst.length}` : ''}` })); }
      if (fs_.length) {
        corpo.push(par(run('Registro fotográfico', { b: true, sz: 9.5, cor: NAVY }), { keep: true, esp: 60, antes: 160 }));
        const colW = Math.floor(LARG / 2);
        const cels = [];
        for (const x of fs_) {
          let b = null; try { b = await lerFoto(x.f); } catch (e) { }
          if (!b) { cels.push(celula([par(run('(foto não encontrada no servidor)', { i: true, sz: 8.5 }), { al: 'center', esp: 0 }), par(run(x.legenda, { sz: 8.5 }), { al: 'center', esp: 80 })], colW, { semBorda: true })); continue; }
          const { w: pw, h: ph } = tamanhoJPEG(b);
          let wcm = 8, hcm = wcm * ph / pw; if (hcm > 6.2) { hcm = 6.2; wcm = hcm * pw / ph; }
          const { rid, nome } = addImg(b);
          cels.push(celula([par(imgInline(rid, Math.round(wcm * CM), Math.round(hcm * CM), nome), { al: 'center', esp: 40 }), par(run(x.legenda, { sz: 8.5, cor: '4B5373' }), { al: 'center', esp: 120 })], colW, { semBorda: true }));
        }
        const linhasF = [];
        for (let i = 0; i < cels.length; i += 2) linhasF.push(cels[i] + (cels[i + 1] || celula('', colW, { semBorda: true })));
        corpo.push(tabela(linhasF, [colW, colW], { semBorda: true }));
      }
    }
  }

  /* 4. considerações finais */
  corpo.push(par('4. Considerações finais', { estilo: 'Heading1' }));
  if (os.parecer && String(os.parecer).trim()) corpo.push(par(os.parecer, { keep: true }));
  else corpo.push(par(run('[Escreva aqui o parecer técnico: condição geral da instalação, pontos de atenção e recomendações.]', { hl: 'yellow' }), { keep: true }));

  /* assinaturas */
  corpo.push(par('', { esp: 240, keep: true }));
  const wA = Math.floor(LARG / 2);
  const ass = (titulo, nome) => celula([par(run('______________________________________', { cor: '7A819D' }), { al: 'center', esp: 0 }), par(run(nome || ' ', { b: true, sz: 9.5 }), { al: 'center', esp: 0 }), par(run(titulo, { sz: 9, cor: '4B5373' }), { al: 'center', esp: 0 })], wA, { semBorda: true });
  corpo.push(tabela([ass('Técnico executante', nomesTec), ass('Responsável técnico', '')], [wA, wA], { semBorda: true }));

  /* seção de conteúdo (timbrado) — área livre: topo 94 pt, base 718 pt, margens 55 pt */
  const temTimbrado = !!ler('timbrado.jpg');
  const secConteudo = `<w:headerReference w:type="default" r:id="rId3"/><w:footerReference w:type="default" r:id="rId5"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="${temTimbrado ? 1880 : 1134}" w:right="1100" w:bottom="${temTimbrado ? 2480 : 1134}" w:left="1100" w:header="0" w:footer="${temTimbrado ? 1560 : 500}" w:gutter="0"/><w:pgNumType w:start="1"/>`;

  /* fechamento */
  const fechs = comCapa ? ['fechamento-1.jpg', 'fechamento-2.jpg'].map(ler).filter(Boolean) : [];
  if (fechs.length) {
    corpo.push(`<w:p><w:pPr><w:sectPr>${secConteudo}</w:sectPr></w:pPr></w:p>`);
    fechs.forEach((b, i) => { const { rid, nome } = addImg(b); corpo.push(`<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>${imgFundo(rid, nome)}${i < fechs.length - 1 ? '<w:r><w:br w:type="page"/></w:r>' : ''}</w:p>`); });
  }
  const secFinal = fechs.length ? `<w:headerReference w:type="default" r:id="rId4"/><w:footerReference w:type="default" r:id="rId6"/>${secCapa}` : secConteudo;
  const documento = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${corpo.join('')}<w:sectPr>${secFinal}</w:sectPr></w:body></w:document>`;

  /* cabeçalho com o timbrado por trás da página */
  let header1 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${NS}><w:p><w:pPr><w:spacing w:before="0" w:after="0"/><w:rPr><w:sz w:val="2"/></w:rPr></w:pPr>`;
  let header1Rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
  if (temTimbrado) { midia.push({ nome: 'timbrado.jpg', dados: ler('timbrado.jpg') }); header1 += imgFundo('rIdT', 'timbrado.jpg'); header1Rels += '<Relationship Id="rIdT" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/timbrado.jpg"/>'; }
  header1 += '</w:p></w:hdr>'; header1Rels += '</Relationships>';
  const header2 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${NS}><w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:p></w:hdr>`;
  const fld = c => `<w:r><w:rPr><w:sz w:val="16"/><w:color w:val="4B5373"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:rPr><w:sz w:val="16"/><w:color w:val="4B5373"/></w:rPr><w:instrText xml:space="preserve"> ${c} </w:instrText></w:r><w:r><w:rPr><w:sz w:val="16"/><w:color w:val="4B5373"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:sz w:val="16"/><w:color w:val="4B5373"/></w:rPr><w:t>1</w:t></w:r><w:r><w:rPr><w:sz w:val="16"/><w:color w:val="4B5373"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>`;
  const footer1 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ${NS}><w:p><w:pPr><w:jc w:val="right"/><w:spacing w:before="0" w:after="0"/></w:pPr>${run(`OS ${os.numero || ''} · página `, { sz: 8, cor: '4B5373' })}${fld('PAGE')}</w:p></w:ftr>`;
  const footer2 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ${NS}><w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:p></w:ftr>`;

  const estilos = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="21"/><w:szCs w:val="21"/><w:color w:val="141A3A"/><w:lang w:val="pt-BR"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="100" w:line="264" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="360" w:after="120"/><w:pBdr><w:bottom w:val="single" w:sz="8" w:space="2" w:color="${GREEN}"/></w:pBdr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:color w:val="${NAVY}"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="280" w:after="100"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:color w:val="${NAVY}"/><w:sz w:val="25"/><w:szCs w:val="25"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="220" w:after="80"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:b/><w:color w:val="${GREEN}"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
</w:styles>`;
  const settings = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:updateFields w:val="false"/><w:defaultTabStop w:val="708"/><w:characterSpacingControl w:val="doNotCompress"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`;
  const relsDocXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/><Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer2.xml"/>${relsDoc.join('')}</Relationships>`;
  const tipos = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="jpg" ContentType="image/jpeg"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/header2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/><Override PartName="/word/footer2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>`;
  const relsRaiz = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`;
  const core = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${esc(`Relatório OS ${os.numero || ''} - ${os.titulo || ''}`)}</dc:title><dc:creator>${esc(empresa || 'Cogera')}</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${new Date().toISOString().slice(0, 19)}Z</dcterms:created></cp:coreProperties>`;

  return zip([
    { nome: '[Content_Types].xml', dados: tipos },
    { nome: '_rels/.rels', dados: relsRaiz },
    { nome: 'docProps/core.xml', dados: core },
    { nome: 'word/document.xml', dados: documento },
    { nome: 'word/styles.xml', dados: estilos },
    { nome: 'word/settings.xml', dados: settings },
    { nome: 'word/header1.xml', dados: header1 },
    { nome: 'word/_rels/header1.xml.rels', dados: header1Rels },
    { nome: 'word/header2.xml', dados: header2 },
    { nome: 'word/footer1.xml', dados: footer1 },
    { nome: 'word/footer2.xml', dados: footer2 },
    { nome: 'word/_rels/document.xml.rels', dados: relsDocXml },
    ...midia.map(m => ({ nome: 'word/media/' + m.nome, dados: m.dados })),
  ]);
}

window.gerarRelatorioDocx = gerarRelatorio;
})();
