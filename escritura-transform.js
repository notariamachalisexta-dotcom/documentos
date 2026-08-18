// Transformación determinista borrador bancario -> escritura pública notarial (OOXML directo)
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XMLNS = 'http://www.w3.org/XML/1998/namespace';

const TITULO_RE = /^(COMPRAVENTA|MUTUO|HIPOTECA|CANCELACI[ÓO]N|ALZAMIENTO)/i;
const SEP_RE = /^[\*\-_=\.]{3,}$/;
// patrones exclusivos de la carátula interna del banco (etiqueta + separador)
const CARATULA_PATTERNS = [
  { label: 'CTA GASTO', re: /CTA\.?\s*GASTO/ },
  { label: 'SUCURSAL :', re: /SUCURSAL\s*[:\.]/ },
  { label: 'EJECUTIVO :', re: /EJECUTIVO\s*[:\.]/ },
  { label: 'PROTOCOLIZAR', re: /PROTOCOLIZAR/ },
  { label: 'INSTRUCCIONES A NOTARIA', re: /INSTRUCCIONES\s+A\s+NOTAR[IÍ]A/ },
  { label: 'Departamento Legal Hipotecario', re: /Departamento\s+Legal\s+Hipotecario/i },
  { label: 'FORMULARIO B', re: /FORMULARIO\s+B/ },
];

function el(doc, name) { return doc.createElementNS(W, name); }
function attr(node, name, val) { node.setAttributeNS(W, name, val); }

function textOf(node) {
  let s = '';
  const walk = (n) => {
    for (const c of Array.from(n.childNodes)) {
      if (c.nodeType !== 1) continue;
      const tn = c.nodeName;
      if (tn === 'w:t') s += c.textContent;
      else if (tn === 'w:tab') s += '\t';
      else if (tn === 'w:br') s += '\n';
      else walk(c);
    }
  };
  walk(node);
  return s;
}

function norm(s) { return s.replace(/\s+/g, ' ').trim(); }

function bodyParagraphs(body) {
  return Array.from(body.childNodes).filter(n => n.nodeType === 1 && n.nodeName === 'w:p');
}

function isCentered(p) {
  const jc = p.getElementsByTagName('w:jc')[0];
  return jc && jc.getAttributeNS(W, 'val') === 'center';
}
function hasBold(p) { return p.getElementsByTagName('w:b').length > 0; }

function makeRun(doc, text, opts = {}) {
  const r = el(doc, 'w:r');
  const rPr = el(doc, 'w:rPr');
  const fonts = el(doc, 'w:rFonts');
  attr(fonts, 'w:ascii', 'Arial'); attr(fonts, 'w:hAnsi', 'Arial'); attr(fonts, 'w:cs', 'Arial');
  rPr.appendChild(fonts);
  if (opts.bold) { rPr.appendChild(el(doc, 'w:b')); rPr.appendChild(el(doc, 'w:bCs')); }
  if (opts.underline) { const u = el(doc, 'w:u'); attr(u, 'w:val', 'single'); rPr.appendChild(u); }
  if (opts.spacing !== undefined) { const sp = el(doc, 'w:spacing'); attr(sp, 'w:val', String(opts.spacing)); rPr.appendChild(sp); }
  if (opts.highlight) { const h = el(doc, 'w:highlight'); attr(h, 'w:val', opts.highlight); rPr.appendChild(h); }
  r.appendChild(rPr);
  if (opts.tab) r.appendChild(el(doc, 'w:tab'));
  if (text !== undefined && text !== null) {
    const t = el(doc, 'w:t');
    t.setAttributeNS(XMLNS, 'xml:space', 'preserve');
    t.textContent = text;
    r.appendChild(t);
  }
  return r;
}

function makePara(doc, runs, opts = {}) {
  const p = el(doc, 'w:p');
  const pPr = el(doc, 'w:pPr');
  const sp = el(doc, 'w:spacing');
  attr(sp, 'w:line', '360'); attr(sp, 'w:lineRule', 'auto');
  attr(sp, 'w:before', '0'); attr(sp, 'w:after', '0');
  pPr.appendChild(sp);
  const jc = el(doc, 'w:jc'); attr(jc, 'w:val', opts.jc || 'both'); pPr.appendChild(jc);
  const ta = el(doc, 'w:textAlignment'); attr(ta, 'w:val', 'baseline'); pPr.appendChild(ta);
  p.appendChild(pPr);
  (runs || []).forEach(r => p.appendChild(r));
  return p;
}

function emptyParas(doc, n, jc) {
  const out = [];
  for (let i = 0; i < n; i++) out.push(makePara(doc, [], { jc: jc || 'both' }));
  return out;
}

// ---------- análisis de la carátula ----------
function parseZipXml(xmlStr) {
  return new DOMParser().parseFromString(xmlStr, 'application/xml');
}

function findAnchorIndex(paras) {
  let fallback = -1;
  for (let i = 0; i < paras.length; i++) {
    const t = norm(textOf(paras[i]));
    if (!TITULO_RE.test(t)) continue;
    if (t.length > 120) continue;
    if (isCentered(paras[i]) && hasBold(paras[i])) return i;
    if (fallback === -1) fallback = i;
  }
  return fallback;
}

export function analyzeXml(docXml) {
  const doc = parseZipXml(docXml);
  const body = doc.getElementsByTagName('w:body')[0];
  const paras = bodyParagraphs(body);
  const anchor = findAnchorIndex(paras);
  const head = anchor > 0 ? paras.slice(0, anchor) : [];

  const instrucciones = {};
  const protocolizar = [];
  let inProto = false;
  const labelRe = /^(CLIENTE|ABOGADO|SUCURSAL|EJECUTIVO|NOTAR[IÍ]A|CTA\.?\s*GASTO(?:\s*N°)?|MES QUE DEBE TENER LA ESCRITURA)\s*[:.\-]\s*(.+)$/i;
  for (const p of head) {
    const t = norm(textOf(p));
    if (!t) continue;
    if (/PROTOCOLIZAR/i.test(t)) { inProto = true; continue; }
    const m = t.match(labelRe);
    if (m) {
      const key = m[1].toUpperCase().replace(/\s+/g, '_').replace(/[.°]/g, '');
      instrucciones[key] = m[2].trim();
      inProto = false;
      continue;
    }
    if (inProto) protocolizar.push(t.replace(/^\d+[).\-]?\s*/, ''));
  }

  // sugerencias de carátula B2 (entre el título y "comparecen")
  const sug = { titulo_acto: '', comprador: '', vendedor: '', banco: '', proyecto: '', subsidio: '' };
  if (anchor >= 0) {
    sug.titulo_acto = norm(textOf(paras[anchor]));
    const rest = [];
    for (let i = anchor + 1; i < paras.length; i++) {
      const t = norm(textOf(paras[i]));
      if (/comparecen/i.test(t)) break;
      if (t && !/^(A|Y)$/i.test(t)) rest.push(t);
      if (rest.length > 8) break;
    }
    for (const t of rest) {
      if (/D\.?\s*S\.?\s*N/i.test(t) && !sug.subsidio) sug.subsidio = t;
      else if (/BANCO|SANTANDER|SCOTIABANK|BICE|ITA[UÚ]|FALABELLA|SECURITY|CHILE\b/i.test(t) && !sug.banco) sug.banco = t;
      else if (!sug.comprador) sug.comprador = t;
      else if (!sug.vendedor) sug.vendedor = t;
      else if (!sug.proyecto) sug.proyecto = t;
    }
  }
  return { instrucciones, protocolizar, sugerencias: sug, anchorFound: anchor >= 0, totalParrafos: paras.length };
}

// ---------- pipeline ----------
function preambuloRuns(doc, data) {
  const n = data.notaria, f = data.fecha;
  const dia = (f.dia && String(f.dia).trim()) ? String(f.dia).trim() : null;
  const runs = [];
  runs.push(makeRun(doc, `En la comuna de ${n.comuna}, Provincia de ${n.provincia}, República de Chile, a `));
  if (dia) runs.push(makeRun(doc, dia, { bold: true }));
  else runs.push(makeRun(doc, '******', { bold: true, highlight: 'yellow' }));
  runs.push(makeRun(doc, ' de '));
  runs.push(makeRun(doc, f.mes, { bold: true }));
  runs.push(makeRun(doc, ' del año '));
  runs.push(makeRun(doc, f.anio_palabras, { bold: true }));
  runs.push(makeRun(doc, ', ante mí, '));
  runs.push(makeRun(doc, n.notario, { bold: true }));
  runs.push(makeRun(doc, `, ${n.titulo}, ${n.cargo} de la ${n.oficio}, según acta ${n.acta} y Decreto de Nombramiento número ${n.decreto}, que se encuentra protocolizado bajo el repertorio ${n.repertorio_nombramiento}, con oficio en esta ciudad, ${n.domicilio}, comparecen: `));
  return runs;
}

function cierreRuns(doc) {
  return [
    makeRun(doc, ' En comprobante y previa lectura y ratificación, firman los comparecientes. La presente escritura se anotó en el Repertorio a mi cargo. '),
    makeRun(doc, 'DOY FE.', { bold: true }),
  ];
}

function contentChildren(p) {
  return Array.from(p.childNodes).filter(n => !(n.nodeType === 1 && n.nodeName === 'w:pPr'));
}

function lastRunRPr(p) {
  const runs = p.getElementsByTagName('w:r');
  if (!runs.length) return null;
  const rPr = runs[runs.length - 1].getElementsByTagName('w:rPr')[0];
  return rPr ? rPr.cloneNode(true) : null;
}

function collapseSpaces(para) {
  const ts = Array.from(para.getElementsByTagName('w:t'));
  let prevSpace = true;
  for (const t of ts) {
    let s = t.textContent.replace(/[\t\n\r]+/g, ' ').replace(/ {2,}/g, ' ');
    if (prevSpace) s = s.replace(/^ +/, '');
    if (s === '' ) { t.textContent = ''; continue; }
    prevSpace = / $/.test(s);
    t.setAttributeNS(XMLNS, 'xml:space', 'preserve');
    t.textContent = s;
  }
}

function setSectPr(doc, body) {
  let sect = null;
  for (const c of Array.from(body.childNodes)) if (c.nodeType === 1 && c.nodeName === 'w:sectPr') sect = c;
  if (!sect) { sect = el(doc, 'w:sectPr'); body.appendChild(sect); }
  // limpiar referencias a header/footer y elementos de página
  for (const tag of ['w:headerReference', 'w:footerReference', 'w:pgSz', 'w:pgMar', 'w:cols', 'w:docGrid', 'w:titlePg']) {
    Array.from(sect.getElementsByTagName(tag)).forEach(n => n.parentNode.removeChild(n));
  }
  const pgSz = el(doc, 'w:pgSz');
  attr(pgSz, 'w:w', '12240'); attr(pgSz, 'w:h', '18720'); attr(pgSz, 'w:code', '281');
  const pgMar = el(doc, 'w:pgMar');
  attr(pgMar, 'w:top', '2268'); attr(pgMar, 'w:right', '2268'); attr(pgMar, 'w:bottom', '3969');
  attr(pgMar, 'w:left', '2268'); attr(pgMar, 'w:header', '709'); attr(pgMar, 'w:footer', '709'); attr(pgMar, 'w:gutter', '0');
  const cols = el(doc, 'w:cols'); attr(cols, 'w:space', '720');
  const grid = el(doc, 'w:docGrid'); attr(grid, 'w:linePitch', '360');
  sect.appendChild(pgSz); sect.appendChild(pgMar); sect.appendChild(cols); sect.appendChild(grid);
  return sect;
}

function normalizeRun(doc, r) {
  let rPr = r.getElementsByTagName('w:rPr')[0];
  if (!rPr || rPr.parentNode !== r) {
    rPr = el(doc, 'w:rPr');
    r.insertBefore(rPr, r.firstChild);
  }
  Array.from(rPr.getElementsByTagName('w:sz')).forEach(n => n.parentNode.removeChild(n));
  Array.from(rPr.getElementsByTagName('w:szCs')).forEach(n => n.parentNode.removeChild(n));
  let fonts = rPr.getElementsByTagName('w:rFonts')[0];
  if (!fonts) { fonts = el(doc, 'w:rFonts'); rPr.insertBefore(fonts, rPr.firstChild); }
  attr(fonts, 'w:ascii', 'Arial'); attr(fonts, 'w:hAnsi', 'Arial'); attr(fonts, 'w:cs', 'Arial');
  fonts.removeAttributeNS(W, 'asciiTheme'); fonts.removeAttributeNS(W, 'hAnsiTheme'); fonts.removeAttributeNS(W, 'cstheme');
}

function normalizePara(doc, p, jcVal) {
  let pPr = null;
  for (const c of Array.from(p.childNodes)) if (c.nodeType === 1 && c.nodeName === 'w:pPr') pPr = c;
  const keepStyle = pPr ? pPr.getElementsByTagName('w:pStyle')[0] : null;
  let markRPr = null, innerSect = null;
  if (pPr) {
    for (const c of Array.from(pPr.childNodes)) {
      if (c.nodeType !== 1) continue;
      if (c.nodeName === 'w:rPr') markRPr = c;
      if (c.nodeName === 'w:sectPr') innerSect = c;
    }
  }
  const newPPr = el(doc, 'w:pPr');
  if (keepStyle) newPPr.appendChild(keepStyle.cloneNode(true));
  const sp = el(doc, 'w:spacing');
  attr(sp, 'w:line', '360'); attr(sp, 'w:lineRule', 'auto'); attr(sp, 'w:before', '0'); attr(sp, 'w:after', '0');
  newPPr.appendChild(sp);
  const jc = el(doc, 'w:jc'); attr(jc, 'w:val', jcVal); newPPr.appendChild(jc);
  const ta = el(doc, 'w:textAlignment'); attr(ta, 'w:val', 'baseline'); newPPr.appendChild(ta);
  if (markRPr) {
    const clone = markRPr.cloneNode(true);
    Array.from(clone.getElementsByTagName('w:sz')).forEach(n => n.parentNode.removeChild(n));
    Array.from(clone.getElementsByTagName('w:szCs')).forEach(n => n.parentNode.removeChild(n));
    let fonts = clone.getElementsByTagName('w:rFonts')[0];
    if (!fonts) { fonts = el(doc, 'w:rFonts'); clone.insertBefore(fonts, clone.firstChild); }
    attr(fonts, 'w:ascii', 'Arial'); attr(fonts, 'w:hAnsi', 'Arial'); attr(fonts, 'w:cs', 'Arial');
    newPPr.appendChild(clone);
  }
  if (innerSect) newPPr.appendChild(innerSect.cloneNode(true));
  if (pPr) p.replaceChild(newPPr, pPr); else p.insertBefore(newPPr, p.firstChild);
  Array.from(p.getElementsByTagName('w:r')).forEach(r => normalizeRun(doc, r));
}

function transformDocumentXml(docXml, data) {
  const doc = parseZipXml(docXml);
  const body = doc.getElementsByTagName('w:body')[0];
  let paras = bodyParagraphs(body);

  // PASO 2 — recortar carátula
  const anchor = findAnchorIndex(paras);
  if (anchor > 0) for (let i = 0; i < anchor; i++) paras[i].parentNode.removeChild(paras[i]);

  // PASO 3 — separadores
  paras = bodyParagraphs(body);
  paras.forEach(p => { if (SEP_RE.test(norm(textOf(p)).replace(/\s/g, ''))) p.parentNode.removeChild(p); });

  // PASO 5 — preámbulo notarial antes de "comparecen"
  paras = bodyParagraphs(body);
  let compIdx = paras.findIndex(p => /comparecen/i.test(textOf(p)));
  if (compIdx === -1) compIdx = Math.min(paras.length - 1, 8);
  const compPara = paras[compIdx];
  // recortar dentro del párrafo todo lo anterior a "comparecen:" inclusive
  const kids = contentChildren(compPara);
  let found = false;
  for (const k of kids) {
    if (found) break;
    const txt = textOf(k);
    if (/comparecen/i.test(txt)) {
      found = true;
      const tNodes = Array.from(k.getElementsByTagName ? k.getElementsByTagName('w:t') : []);
      for (const t of tNodes) {
        const m = t.textContent.match(/comparecen\s*:?/i);
        if (m) {
          t.textContent = t.textContent.slice(m.index + m[0].length);
          t.setAttributeNS(XMLNS, 'xml:space', 'preserve');
          break;
        } else t.textContent = '';
      }
    } else {
      k.parentNode.removeChild(k);
    }
  }
  preambuloRuns(doc, data).reverse().forEach(r => {
    compPara.insertBefore(r, contentChildren(compPara)[0] || null);
  });

  // PASO 6 — fundir el cuerpo en un solo párrafo
  paras = bodyParagraphs(body);
  const startIdx = paras.indexOf(compPara);
  let endIdx = paras.length - 1;
  while (endIdx > startIdx && !norm(textOf(paras[endIdx]))) endIdx--;

  const merged = makePara(doc, [], { jc: 'both' });
  let acc = '';
  for (let i = startIdx; i <= endIdx; i++) {
    const p = paras[i];
    const t = textOf(p);
    if (!norm(t)) continue;
    if (acc && !/\s$/.test(acc)) {
      const spRun = makeRun(doc, ' ');
      const prevRPr = lastRunRPr(p);
      if (prevRPr) { const old = spRun.getElementsByTagName('w:rPr')[0]; spRun.replaceChild(prevRPr, old); }
      merged.appendChild(spRun);
      acc += ' ';
    }
    contentChildren(p).forEach(c => merged.appendChild(c.cloneNode(true)));
    acc += t;
  }
  // limpiar tabs, saltos y numeración dentro del cuerpo
  Array.from(merged.getElementsByTagName('w:tab')).forEach(n => n.parentNode.removeChild(n));
  Array.from(merged.getElementsByTagName('w:br')).forEach(n => n.parentNode.removeChild(n));
  Array.from(merged.getElementsByTagName('w:numPr')).forEach(n => n.parentNode.removeChild(n));
  collapseSpaces(merged);

  // PASO 7 — cierre
  if (!/DOY\s+FE/i.test(textOf(merged))) cierreRuns(doc).forEach(r => merged.appendChild(r));

  // reemplazar rango por el párrafo único y borrar lo que sobra tras él
  body.insertBefore(merged, paras[startIdx]);
  for (let i = startIdx; i < paras.length; i++) if (paras[i].parentNode) paras[i].parentNode.removeChild(paras[i]);

  // PASO 7b — firmas
  const firmas = [];
  emptyParas(doc, 4).forEach(p => firmas.push(p));
  const comps = data.comparecientes.filter(c => c.nombre && c.nombre.trim());
  comps.forEach((c, i) => {
    firmas.push(makePara(doc, [makeRun(doc, `${i + 1}.`, { bold: true })], { jc: 'both' }));
    firmas.push(makePara(doc, [makeRun(doc, c.nombre.trim(), { bold: true })], { jc: 'both' }));
    firmas.push(makePara(doc, [makeRun(doc, 'C.I.N°', { bold: true })], { jc: 'both' }));
    if (i < comps.length - 1) emptyParas(doc, 3).forEach(p => firmas.push(p));
  });

  // PASO 4 — repertorio al inicio
  const anio = data.operacion.anio;
  const repRuns = [
    makeRun(doc, 'REPERTORIO ', { bold: true, spacing: -3 }),
    makeRun(doc, 'N°', { bold: true, spacing: -3 }),
    makeRun(doc, null, { bold: true, spacing: -3, tab: true }),
    makeRun(doc, `-${anio}`, { bold: true, spacing: -3 }),
  ];
  const head = [
    makePara(doc, repRuns, { jc: 'both' }),
    makePara(doc, [makeRun(doc, 'OT.', { bold: true, spacing: -3 })], { jc: 'both' }),
    makePara(doc, [], { jc: 'both' }),
  ];

  const sect = setSectPr(doc, body);
  // colapsar los párrafos vacíos previos al cuerpo a exactamente uno
  const beforeBody = bodyParagraphs(body);
  const mergedAt = beforeBody.indexOf(merged);
  let k = mergedAt - 1;
  const emptiesBefore = [];
  while (k >= 0 && !norm(textOf(beforeBody[k]))) { emptiesBefore.push(beforeBody[k]); k--; }
  emptiesBefore.slice(1).forEach(p => p.parentNode.removeChild(p));

  const firstNode = bodyParagraphs(body)[0] || sect;
  head.forEach(p => body.insertBefore(p, firstNode));
  firmas.forEach(p => body.insertBefore(p, sect));

  // PASO 9 — re-maquetar todo
  const all = bodyParagraphs(body);
  const mergedPos = all.indexOf(merged);
  all.forEach((p, i) => {
    const isCaratula = i > 2 && i < mergedPos; // bloque B2 entre repertorio y cuerpo
    normalizePara(doc, p, isCaratula ? 'center' : 'both');
    Array.from(p.getElementsByTagName('w:ind')).forEach(n => n.parentNode.removeChild(n));
    Array.from(p.getElementsByTagName('w:numPr')).forEach(n => n.parentNode.removeChild(n));
    Array.from(p.getElementsByTagName('w:tabs')).forEach(n => n.parentNode.removeChild(n));
    Array.from(p.getElementsByTagName('w:pBdr')).forEach(n => n.parentNode.removeChild(n));
    Array.from(p.getElementsByTagName('w:contextualSpacing')).forEach(n => n.parentNode.removeChild(n));
  });

  return { xml: new XMLSerializer().serializeToString(doc), mergedText: textOf(merged), paraCount: all.length };
}

function patchSettings(xml) {
  let out = xml;
  if (!/mirrorMargins/.test(out)) {
    const mm = '<w:mirrorMargins/>';
    if (/<w:proofState[^>]*\/>/.test(out)) out = out.replace(/(<w:proofState[^>]*\/>)/, mm + '$1');
    else if (/<w:zoom[^>]*(\/>|<\/w:zoom>)/.test(out)) out = out.replace(/(<w:zoom[^>]*(?:\/>|<\/w:zoom>))/, '$1' + mm);
    else out = out.replace(/(<\/w:settings>)/, mm + '$1');
  }
  if (!/hyphenationZone/.test(out)) {
    const hz = '<w:hyphenationZone w:val="425"/>';
    if (/<w:defaultTabStop[^>]*\/>/.test(out)) out = out.replace(/(<w:defaultTabStop[^>]*\/>)/, '$1' + hz);
    else out = out.replace(/(<\/w:settings>)/, hz + '$1');
  }
  out = out.replace(/<w:autoHyphenation[^>]*\/>/g, '');
  return out;
}

function validate(finalXml, mergedText, comps) {
  const v = [];
  const add = (id, label, ok, detail) => v.push({ id, label, ok, detail: detail || '' });
  add('V1', 'Hoja oficio 21,59 × 33,02 cm', /w:w="12240"[^>]*w:h="18720"/.test(finalXml));
  add('V2', 'Márgenes 4/4/4/7 cm y header/footer 1,25 cm', /w:top="2268"[^>]*w:right="2268"[^>]*w:bottom="3969"[^>]*w:left="2268"[^>]*w:header="709"[^>]*w:footer="709"/.test(finalXml));
  const runs = finalXml.match(/<w:rFonts[^>]*>/g) || [];
  const badFont = runs.filter(r => !/w:ascii="Arial"/.test(r)).length;
  add('V3', 'Todos los runs en Arial', badFont === 0, badFont ? `${badFont} runs con otra fuente` : '');
  const szCount = (finalXml.match(/<w:sz\s/g) || []).length;
  add('V4', 'Sin tamaños explícitos (hereda 12 pt)', szCount === 0, szCount ? `${szCount} runs con w:sz` : '');
  const paraCount = (finalXml.match(/<w:p[\s>]/g) || []).length;
  const spacingCount = (finalXml.match(/w:line="360"/g) || []).length;
  add('V5', 'Interlineado 1,5 en todos los párrafos', spacingCount >= paraCount - 1, `${spacingCount}/${paraCount}`);
  const numPr = (finalXml.match(/<w:numPr>/g) || []).length;
  add('V6', 'Sin numeración automática', numPr === 0);
  add('V8', 'Cuerpo sin párrafos vacíos internos', !/\s{3,}$/.test(mergedText));
  add('V9', 'Contiene "ante mí", "comparecen:" y "DOY FE."', /ante mí/i.test(mergedText) && /comparecen/i.test(mergedText) && /DOY FE\./.test(mergedText));
  const ci = (finalXml.match(/C\.I\.N°/g) || []).length;
  add('V10', 'Bloques de firma por compareciente', ci === comps.length, `${ci} de ${comps.length}`);
  const plain = finalXml.replace(/<[^>]+>/g, '');
  const leftovers = CARATULA_PATTERNS.filter(p => p.re.test(plain)).map(p => p.label);
  add('V12', 'Sin restos de la carátula del banco', leftovers.length === 0, leftovers.join(', '));
  const stars = (mergedText.match(/\*{3,}/g) || []).length;
  add('V13', 'Sin separadores de asteriscos (salvo día pendiente)', stars <= 1);
  add('V14', 'Sin dobles espacios ni puntos pegados', !/ {2}/.test(mergedText) && !/\.[A-ZÁÉÍÓÚÑ]{2}/.test(mergedText));
  add('V16', 'Sin encabezado, pie ni número de página', !/w:headerReference|w:footerReference|PAGE\s/.test(finalXml));
  return v;
}

export async function analyzeDocx(arrayBuffer) {
  const zip = await JSZip.loadAsync(arrayBuffer);
  const entry = zip.file('word/document.xml');
  if (!entry) throw new Error('El archivo no es un .docx válido.');
  return analyzeXml(await entry.async('string'));
}

export async function transformDocx(arrayBuffer, data) {
  const zip = await JSZip.loadAsync(arrayBuffer);
  const entry = zip.file('word/document.xml');
  if (!entry) throw new Error('El archivo no es un .docx válido.');
  const { xml, mergedText, paraCount } = transformDocumentXml(await entry.async('string'), data);
  zip.file('word/document.xml', xml);

  const settings = zip.file('word/settings.xml');
  if (settings) zip.file('word/settings.xml', patchSettings(await settings.async('string')));

  // sin encabezados/pies ni logos externos
  for (const name of Object.keys(zip.files)) {
    if (/word\/(header|footer)\d*\.xml$/.test(name)) zip.remove(name);
    if (/word\/numbering\.xml$/.test(name)) zip.remove(name);
  }
  const relsEntry = zip.file('word/_rels/document.xml.rels');
  if (relsEntry) {
    let rels = await relsEntry.async('string');
    rels = rels.replace(/<Relationship[^>]*Type="[^"]*\/(header|footer|numbering)"[^>]*\/>/g, '');
    zip.file('word/_rels/document.xml.rels', rels);
  }
  const ctEntry = zip.file('[Content_Types].xml');
  if (ctEntry) {
    let ct = await ctEntry.async('string');
    ct = ct.replace(/<Override[^>]*PartName="\/word\/(header|footer)\d*\.xml"[^>]*\/>/g, '');
    ct = ct.replace(/<Override[^>]*PartName="\/word\/numbering\.xml"[^>]*\/>/g, '');
    zip.file('[Content_Types].xml', ct);
  }

  const blob = await zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  return { blob, validations: validate(xml, mergedText, data.comparecientes.filter(c => c.nombre && c.nombre.trim())), paraCount, chars: mergedText.length };
}
