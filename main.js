'use strict';
// HWPX = ZIP + XML (OWPML), DOCX = ZIP + XML (WordprocessingML).
// JSZip으로 압축을 풀고, 브라우저 내장 DOMParser로 XML을 읽는다.

const $ = (id) => document.getElementById(id);
const PX = 1 / 75; // HWPUNIT(1/7200 inch) -> px(96dpi)
const MIME = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml' };

// ---- XML 헬퍼: 접두사(hp:, hh:)에 의존하지 않도록 localName으로 찾는다 ----
const kids = (el, n) => Array.from(el.children).filter((c) => c.localName === n);
const first = (el, n) => (el ? kids(el, n)[0] : undefined);
const all = (el, n) => Array.from(el.getElementsByTagNameNS('*', n));
const num = (v) => Number(v) || 0;

function parseXml(text, label) {
  const doc = new DOMParser().parseFromString(text, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) throw new Error(`${label} 파싱 실패`);
  return doc;
}
const findEntry = (zip, path) => {
  const p = path.replace(/^\.?\//, '').toLowerCase();
  return Object.values(zip.files).find((f) => !f.dir && f.name.toLowerCase() === p);
};
async function readXml(zip, path) {
  const f = findEntry(zip, path);
  if (!f) throw new Error(`${path} 파일이 없습니다`);
  return parseXml(await f.async('string'), path);
}
const baseName = (p) => p.split('/').pop().replace(/\.[^.]+$/, '');

// ---- 1단계: header.xml 의 글자/문단 스타일 ----
function parseHeader(doc) {
  const charPr = new Map(), paraPr = new Map();
  for (const el of all(doc, 'charPr')) {
    const ul = first(el, 'underline'), st = first(el, 'strikeout');
    charPr.set(el.getAttribute('id'), {
      size: num(el.getAttribute('height')) / 100, // 1000 = 10pt
      color: el.getAttribute('textColor'),
      bold: !!first(el, 'bold'),
      italic: !!first(el, 'italic'),
      underline: !!ul && ul.getAttribute('type') !== 'NONE',
      strike: !!st && st.getAttribute('shape') !== 'NONE',
    });
  }
  for (const el of all(doc, 'paraPr')) {
    const align = all(el, 'align')[0]?.getAttribute('horizontal');
    const ls = all(el, 'lineSpacing')[0];
    paraPr.set(el.getAttribute('id'), {
      align: { CENTER: 'center', RIGHT: 'right', JUSTIFY: 'justify', LEFT: 'left' }[align],
      lineHeight: ls && ls.getAttribute('type') === 'PERCENT' ? num(ls.getAttribute('value')) / 100 : 0,
    });
  }
  return { charPr, paraPr };
}

// ---- 2단계: content.hpf (섹션 순서, 이미지 ID -> 파일) ----
async function loadManifest(zip) {
  const f = findEntry(zip, 'Contents/content.hpf');
  const items = new Map(); let spine = [];
  if (f) {
    const doc = parseXml(await f.async('string'), 'content.hpf');
    for (const it of all(doc, 'item')) items.set(it.getAttribute('id'), it.getAttribute('href'));
    spine = all(doc, 'itemref').map((r) => items.get(r.getAttribute('idref'))).filter((h) => /section\d+\.xml$/i.test(h || ''));
  }
  if (!spine.length) {
    spine = Object.keys(zip.files).filter((n) => /^contents\/section\d+\.xml$/i.test(n)).sort();
  }
  return { items, spine };
}
async function loadImages(zip, manifest) {
  const urls = new Map();
  for (const f of Object.values(zip.files)) {
    if (f.dir || !/^bindata\//i.test(f.name)) continue;
    const ext = f.name.split('.').pop().toLowerCase();
    const blob = new Blob([await f.async('uint8array')], { type: MIME[ext] || '' });
    urls.set(baseName(f.name), URL.createObjectURL(blob));
  }
  for (const [id, href] of manifest.items) {
    const u = href && urls.get(baseName(href));
    if (u) urls.set(id, u);
  }
  return urls;
}

// ---- 3단계: 렌더링 (textContent / append(text)만 사용 -> XSS 방지) ----
function renderText(t, span) {
  for (const n of t.childNodes) {
    if (n.nodeType === Node.TEXT_NODE) span.append(n.data);
    else if (n.localName === 'lineBreak') span.append(document.createElement('br'));
    else if (n.localName === 'tab') span.append('\u2003');
    else if (n.localName === 'nbSpace') span.append('\u00A0');
    else if (n.localName === 'fwSpace') span.append('\u3000');
  }
}
function applyChar(span, c) {
  if (!c) return;
  if (c.size) span.style.fontSize = c.size + 'pt';
  if (/^#[0-9a-f]{6}$/i.test(c.color || '')) span.style.color = c.color;
  if (c.bold) span.style.fontWeight = 'bold';
  if (c.italic) span.style.fontStyle = 'italic';
  const deco = [c.underline && 'underline', c.strike && 'line-through'].filter(Boolean).join(' ');
  if (deco) span.style.textDecoration = deco;
}
function missing(msg) {
  const s = document.createElement('span'); s.className = 'missing'; s.textContent = msg; return s;
}
function renderPic(pic, ctx) {
  const ref = all(pic, 'img')[0]?.getAttribute('binaryItemIDRef');
  const url = ctx.urls.get(ref);
  if (!url) return missing(`[이미지 없음: ${ref}]`);
  const w = num(first(pic, 'curSz')?.getAttribute('width')) || num(first(pic, 'sz')?.getAttribute('width'));
  const img = new Image();
  img.src = url; img.alt = '문서 이미지'; img.style.maxWidth = '100%';
  if (w) img.style.width = w * PX + 'px';
  img.onerror = () => img.replaceWith(missing('[표시할 수 없는 이미지 형식]'));
  return img;
}
function renderTable(tbl, ctx) {
  const t = document.createElement('table'); t.className = 'hwpx-tbl';
  const sz = first(tbl, 'sz');
  if (sz) t.style.width = num(sz.getAttribute('width')) * PX + 'px';
  for (const tr of kids(tbl, 'tr')) {
    const row = t.insertRow();
    for (const tc of kids(tr, 'tc')) {
      const td = row.insertCell();
      const span = first(tc, 'cellSpan');
      if (span) { td.colSpan = num(span.getAttribute('colSpan')) || 1; td.rowSpan = num(span.getAttribute('rowSpan')) || 1; }
      const cw = num(first(tc, 'cellSz')?.getAttribute('width'));
      if (cw) td.style.width = cw * PX + 'px';
      for (const p of kids(first(tc, 'subList') || tc, 'p')) td.append(renderParagraph(p, ctx));
    }
  }
  return t;
}
function renderParagraph(p, ctx) {
  const box = document.createElement('div'); box.className = 'hwpx-p';
  const pp = ctx.styles.paraPr.get(p.getAttribute('paraPrIDRef'));
  if (pp?.align) box.style.textAlign = pp.align;
  if (pp?.lineHeight) box.style.lineHeight = pp.lineHeight;
  for (const run of kids(p, 'run')) {
    const cp = ctx.styles.charPr.get(run.getAttribute('charPrIDRef'));
    for (const node of run.children) {
      if (node.localName === 't') {
        const span = document.createElement('span');
        applyChar(span, cp); renderText(node, span); box.append(span);
      } else if (node.localName === 'tbl') box.append(renderTable(node, ctx));
      else if (node.localName === 'pic') box.append(renderPic(node, ctx));
    }
  }
  if (!box.childNodes.length) box.append(document.createElement('br')); // 빈 문단도 줄 높이 유지
  return box;
}
function renderSection(doc, ctx) {
  const sheet = document.createElement('section'); sheet.className = 'sheet';
  const pg = all(doc, 'pagePr')[0], mg = first(pg, 'margin');
  if (pg) sheet.style.width = num(pg.getAttribute('width')) * PX + 'px';
  if (mg) sheet.style.padding = ['top', 'right', 'bottom', 'left'].map((k) => num(mg.getAttribute(k)) * PX + 'px').join(' ');
  for (const p of kids(doc.documentElement, 'p')) sheet.append(renderParagraph(p, ctx));
  return sheet;
}

// ======================================================================
// DOCX = ZIP + XML (WordprocessingML). HWPX와 같은 방식(JSZip + DOMParser)으로 읽는다.
// 지원: 문단/글자 서식(스타일 상속), 표(병합·테두리·음영), 이미지, 하이퍼링크,
//       목록 번호, 각주/미주, 쪽 크기·여백. (머리글/바닥글, 도형(VML/DrawingML)은 생략)
// ======================================================================
const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const wa = (el, n) => (el ? el.getAttributeNS(NS_W, n) : null);          // w:val 같은 속성
const nn = (v) => (v == null || v === '' ? undefined : Number(v));        // 없으면 undefined
const tw = (v) => (Number(v) || 0) / 15;                                   // twip -> px (96dpi)
const pt2px = (v) => ((parseFloat(v) || 0) * 96) / 72;
const hex = (v) => (/^[0-9a-f]{6}$/i.test(v || '') ? '#' + v : undefined);
const mergeProps = (...os) => Object.assign({}, ...os.map((o) => Object.fromEntries(Object.entries(o || {}).filter(([, v]) => v !== undefined))));
function onOff(el) { if (!el) return undefined; const v = wa(el, 'val'); return !(v === '0' || v === 'false' || v === 'off'); }

// ---- 1단계: 글자/문단/표 속성 파싱 ----
const HIGHLIGHT = { yellow: '#ffff00', green: '#00ff00', cyan: '#00ffff', magenta: '#ff00ff', blue: '#0000ff', red: '#ff0000', darkBlue: '#000080', darkCyan: '#008080', darkGreen: '#008000', darkMagenta: '#800080', darkRed: '#800000', darkYellow: '#808000', darkGray: '#808080', lightGray: '#c0c0c0', black: '#000000', white: '#ffffff' };
function parseRPr(el) {
  const r = {}; if (!el) return r;
  const rf = first(el, 'rFonts');
  if (rf) { r.font = wa(rf, 'ascii') || wa(rf, 'hAnsi') || wa(rf, 'eastAsia') || undefined; r.fontEa = wa(rf, 'eastAsia') || undefined; }
  const sz = first(el, 'sz'); if (sz) r.size = Number(wa(sz, 'val')) / 2;           // half-point -> pt
  for (const [k, n] of [['bold', 'b'], ['italic', 'i'], ['strike', 'strike'], ['caps', 'caps'], ['smallCaps', 'smallCaps']]) r[k] = onOff(first(el, n));
  const u = first(el, 'u'); if (u) r.underline = wa(u, 'val') !== 'none';
  const c = first(el, 'color'); if (c) r.color = wa(c, 'val') === 'auto' ? '' : hex(wa(c, 'val'));
  const hl = first(el, 'highlight'); if (hl) r.fill = HIGHLIGHT[wa(hl, 'val')];
  const sh = first(el, 'shd'); if (sh && !hl) r.fill = hex(wa(sh, 'fill'));
  const va = first(el, 'vertAlign'); if (va) r.vert = wa(va, 'val');
  const sp = first(el, 'spacing'); if (sp) r.spacing = nn(wa(sp, 'val'));          // 글자 간격(1/20pt)
  return r;
}
function parsePPr(el) {
  const o = {}; if (!el) return o;
  o.jc = wa(first(el, 'jc'), 'val') || undefined;
  const ind = first(el, 'ind');
  if (ind) {
    const g = (a, b) => nn(wa(ind, a) ?? wa(ind, b));
    o.left = g('left', 'start'); o.right = g('right', 'end'); o.firstLine = g('firstLine', 'firstLine'); o.hanging = g('hanging', 'hanging');
    if (o.firstLine !== undefined && o.hanging === undefined) o.hanging = 0;
    if (o.hanging !== undefined && o.firstLine === undefined) o.firstLine = 0;
  }
  const sp = first(el, 'spacing');
  if (sp) { o.before = nn(wa(sp, 'before')); o.after = nn(wa(sp, 'after')); o.line = nn(wa(sp, 'line')); o.lineRule = wa(sp, 'lineRule') || undefined; }
  const sh = first(el, 'shd'); if (sh) o.fill = hex(wa(sh, 'fill'));
  const np = first(el, 'numPr'); if (np) o.numPr = { id: wa(first(np, 'numId'), 'val'), lvl: Number(wa(first(np, 'ilvl'), 'val')) || 0 };
  if (first(el, 'pageBreakBefore')) o.pageBreakBefore = onOff(first(el, 'pageBreakBefore'));
  return o;
}
function parseBorders(el) {
  const b = {}; if (!el) return b;
  for (const k of ['top', 'left', 'bottom', 'right', 'start', 'end', 'insideH', 'insideV']) {
    const e = first(el, k);
    if (e) b[{ start: 'left', end: 'right' }[k] || k] = { val: wa(e, 'val'), sz: Number(wa(e, 'sz')) || 4, color: wa(e, 'color') };
  }
  return b;
}
function cssBorder(b) {
  if (!b || !b.val || b.val === 'nil' || b.val === 'none') return 'none';
  const w = Math.max(1, Math.round(((b.sz / 8) * 96) / 72));                       // 1/8pt -> px
  const style = { double: 'double', dotted: 'dotted', dashed: 'dashed', dashSmallGap: 'dashed' }[b.val] || 'solid';
  return `${style === 'double' ? Math.max(3, w) : w}px ${style} ${hex(b.color) || '#000'}`;
}

// ---- 2단계: styles.xml / numbering.xml / rels / 각주 ----
function parseStyles(doc) {
  const s = { defaults: { p: {}, r: {} }, byId: new Map(), defPara: null };
  const dd = all(doc, 'docDefaults')[0];
  if (dd) { s.defaults.r = parseRPr(first(first(dd, 'rPrDefault'), 'rPr')); s.defaults.p = parsePPr(first(first(dd, 'pPrDefault'), 'pPr')); }
  for (const st of all(doc, 'style')) {
    const id = wa(st, 'styleId'), type = wa(st, 'type');
    s.byId.set(id, {
      id, type, name: (wa(first(st, 'name'), 'val') || '').toLowerCase(), basedOn: wa(first(st, 'basedOn'), 'val'),
      p: parsePPr(first(st, 'pPr')), r: parseRPr(first(st, 'rPr')), borders: parseBorders(first(first(st, 'tblPr'), 'tblBorders')),
    });
    if (type === 'paragraph' && wa(st, 'default') === '1') s.defPara = id;
  }
  return s;
}
function styleChain(styles, id) {                                                  // 기반 스타일부터 순서대로
  const chain = [], seen = new Set();
  while (id && styles.byId.has(id) && !seen.has(id)) { seen.add(id); chain.unshift(styles.byId.get(id)); id = styles.byId.get(id).basedOn; }
  return chain;
}
function parseNumbering(doc) {
  const abs = new Map(), nums = new Map();
  for (const a of kids(doc.documentElement, 'abstractNum')) {
    const lv = new Map();
    for (const l of kids(a, 'lvl')) {
      lv.set(Number(wa(l, 'ilvl')) || 0, {
        fmt: wa(first(l, 'numFmt'), 'val'), text: wa(first(l, 'lvlText'), 'val') || '',
        start: Number(wa(first(l, 'start'), 'val')) || 1, p: parsePPr(first(l, 'pPr')), r: parseRPr(first(l, 'rPr')),
      });
    }
    abs.set(wa(a, 'abstractNumId'), lv);
  }
  for (const n of kids(doc.documentElement, 'num')) nums.set(wa(n, 'numId'), wa(first(n, 'abstractNumId'), 'val'));
  return { abs, nums };
}
function parseRels(doc) {
  const m = new Map();
  for (const r of all(doc, 'Relationship')) {
    const ext = r.getAttribute('TargetMode') === 'External';
    let t = r.getAttribute('Target') || '';
    if (!ext) t = (t.startsWith('/') ? t.slice(1) : 'word/' + t).replace(/[^/]+\/\.\.\//g, '').toLowerCase();
    m.set(r.getAttribute('Id'), { ext, target: r.getAttribute('Target'), path: t });
  }
  return m;
}
function parseNotes(doc) {
  const m = new Map();
  if (!doc) return m;
  for (const n of kids(doc.documentElement, 'footnote').concat(kids(doc.documentElement, 'endnote'))) {
    if (!/separator/i.test(wa(n, 'type') || '')) m.set(wa(n, 'id'), n);
  }
  return m;
}
async function loadDocxImages(zip) {
  const urls = new Map();
  for (const f of Object.values(zip.files)) {
    if (f.dir || !/^word\/media\//i.test(f.name)) continue;
    const ext = f.name.split('.').pop().toLowerCase();
    urls.set(f.name.toLowerCase(), URL.createObjectURL(new Blob([await f.async('uint8array')], { type: MIME[ext] || '' })));
  }
  return urls;
}

// ---- 3단계: 번호 매기기 ----
const roman = (n) => [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']].reduce((s, [v, c]) => { while (n >= v) { s += c; n -= v; } return s; }, '');
function fmtNum(n, fmt) {
  switch (fmt) {
    case 'lowerLetter': return String.fromCharCode(97 + ((n - 1) % 26));
    case 'upperLetter': return String.fromCharCode(65 + ((n - 1) % 26));
    case 'lowerRoman': return roman(n).toLowerCase();
    case 'upperRoman': return roman(n);
    case 'decimalZero': return String(n).padStart(2, '0');
    case 'decimalEnclosedCircle': return n <= 20 ? String.fromCharCode(0x2460 + n - 1) : String(n);
    case 'ganada': return '가나다라마바사아자차카타파하'[(n - 1) % 14];
    default: return String(n);
  }
}
function nextLabel(ctx, np) {
  if (!np || !np.id || np.id === '0') return null;
  const levels = ctx.numbering.abs.get(ctx.numbering.nums.get(np.id));
  const lv = levels && levels.get(np.lvl);
  if (!lv) return null;
  const cnt = ctx.counters.get(np.id) || [];
  ctx.counters.set(np.id, cnt);
  cnt[np.lvl] = (cnt[np.lvl] ?? lv.start - 1) + 1;
  cnt.length = np.lvl + 1;                                                          // 하위 단계는 다시 1부터
  let label;
  if (lv.fmt === 'bullet') label = !lv.text || /[\uF000-\uF0FF]/.test(lv.text) ? '•' : lv.text;
  else label = lv.text.replace(/%([1-9])/g, (_, k) => { const l = levels.get(k - 1); return l ? fmtNum(cnt[k - 1] ?? l.start, l.fmt) : ''; });
  return { lv, label };
}

// ---- 4단계: 렌더링 (textContent / append(text)만 사용 -> XSS 방지) ----
function applyDocxChar(el, r, text) {
  const s = el.style, size = r.size || 10;
  s.fontSize = size + 'pt';
  if (r.color) s.color = r.color;
  if (r.bold) s.fontWeight = 'bold';
  if (r.italic) s.fontStyle = 'italic';
  const deco = [r.underline && 'underline', r.strike && 'line-through'].filter(Boolean).join(' ');
  if (deco) s.textDecoration = deco;
  const hangul = text && /[\u1100-\u11FF\u3130-\u318F\uAC00-\uD7AF\u4E00-\u9FFF]/.test(text);
  const font = hangul && r.fontEa ? r.fontEa : r.font;
  if (font) s.fontFamily = `"${font.replace(/[^\w \-\u1100-\u11FF\u3130-\u318F\uAC00-\uD7AF]/g, '')}", "Malgun Gothic", sans-serif`;
  if (r.spacing && /\S/.test(text || '')) s.letterSpacing = r.spacing / 20 + 'pt';
  if (r.fill) s.backgroundColor = r.fill;
  if (r.caps) s.textTransform = 'uppercase';
  if (r.smallCaps) s.fontVariant = 'small-caps';
  if (r.vert === 'superscript') { s.verticalAlign = 'super'; s.fontSize = size * 0.7 + 'pt'; }
  if (r.vert === 'subscript') { s.verticalAlign = 'sub'; s.fontSize = size * 0.7 + 'pt'; }
}
function imgEl(url, w, h) {
  const img = new Image();
  img.src = url; img.alt = '문서 이미지';
  if (w) img.style.width = w + 'px';
  if (h) img.style.height = h + 'px';
  img.onerror = () => img.replaceWith(missing('[표시할 수 없는 이미지 형식]'));
  return img;
}
function cssStyleMap(el) {
  return Object.fromEntries(((el && el.getAttribute('style')) || '').split(';').map((s) => s.split(':').map((x) => x.trim())).filter((a) => a.length === 2));
}
function renderDocxDrawing(dr, ctx) {
  const root = dr.firstElementChild;
  const ext = all(dr, 'extent')[0];
  const w = Number(ext && ext.getAttribute('cx')) / 9525, h = Number(ext && ext.getAttribute('cy')) / 9525; // EMU -> px
  const blip = all(dr, 'blip')[0], txb = all(dr, 'txbxContent')[0];
  let node;
  if (blip) {
    const rel = ctx.rels.get(blip.getAttributeNS(NS_R, 'embed') || blip.getAttributeNS(NS_R, 'link'));
    const url = rel && ctx.urls.get(rel.path);
    node = url ? imgEl(url, w, h) : missing('[이미지 없음]');
  } else if (txb) {
    node = document.createElement('div'); node.className = 'docx-textbox';
    if (w) node.style.width = w + 'px';
    const prev = ctx.inCell; ctx.inCell = true; renderBlocks(txb, node, ctx); ctx.inCell = prev;
  } else return null;
  if (root && root.localName === 'anchor' && !ctx.inCell) {                         // 떠 있는 개체: 문단 기준 절대 위치
    const ph = first(root, 'positionH'), pv = first(root, 'positionV');
    const oh = ph && first(ph, 'posOffset'), ov = pv && first(pv, 'posOffset');
    if (oh && ov) {
      const baseH = ph.getAttribute('relativeFrom') === 'page' ? 0 : ctx.page.left;
      node.style.position = 'absolute';
      node.style.left = baseH + Number(oh.textContent) / 9525 - ctx.paraLeftPx + 'px';
      node.style.top = Number(ov.textContent) / 9525 + 'px';
      if (root.getAttribute('behindDoc') === '1') node.style.zIndex = '-1';
    }
  }
  return node;
}
function vmlPathToSvg(d) {                                                          // VML path(m,l,c,x,e) -> SVG path
  const out = []; const re = /([mlcxe])([^mlcxe]*)/gi; let m;
  while ((m = re.exec(d))) {
    const c = m[1].toLowerCase(), v = m[2].split(/[ ,]+/).filter(Boolean).join(' ');
    if (c === 'm') out.push('M' + v); else if (c === 'l') out.push('L' + v); else if (c === 'c') out.push('C' + v); else if (c === 'x') out.push('Z');
  }
  return out.join(' ');
}
function renderVmlGroup(grp, ctx) {
  const st = cssStyleMap(grp);
  const [ox, oy] = (grp.getAttribute('coordorigin') || '0,0').split(',').map(Number);
  const [cw, ch] = (grp.getAttribute('coordsize') || '1,1').split(',').map(Number);
  const gw = pt2px(st.width), gh = pt2px(st.height), sx = gw / (cw || 1), sy = gh / (ch || 1);
  const box = document.createElement('span');
  box.style.cssText = `position:absolute;display:block;left:${pt2px(st['margin-left']) - ctx.paraLeftPx}px;top:${pt2px(st['margin-top'])}px;width:${gw}px;height:${gh}px;`;
  for (const sh of kids(grp, 'shape')) {
    const ss = cssStyleMap(sh), x = (num(ss.left) - ox) * sx, y = (num(ss.top) - oy) * sy, w = num(ss.width) * sx, h = num(ss.height) * sy;
    const im = first(sh, 'imagedata'), d = sh.getAttribute('path');
    let node = null;
    if (im) {
      const rel = ctx.rels.get(im.getAttributeNS(NS_R, 'id')), url = rel && ctx.urls.get(rel.path);
      if (url) node = imgEl(url, w, h);
    } else if (d) {
      const [sox, soy] = (sh.getAttribute('coordorigin') || '0,0').split(',').map(Number);
      const [scw, sch] = (sh.getAttribute('coordsize') || '1,1').split(',').map(Number);
      const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
      svg.setAttribute('viewBox', `${sox} ${soy} ${scw} ${sch}`); svg.setAttribute('width', w); svg.setAttribute('height', h);
      svg.setAttribute('preserveAspectRatio', 'none');
      const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
      path.setAttribute('d', vmlPathToSvg(d)); path.setAttribute('fill-rule', 'evenodd');
      path.setAttribute('fill', hex((sh.getAttribute('fillcolor') || '').slice(1)) || '#000');
      svg.append(path); node = svg;
    }
    if (node) { node.style.position = 'absolute'; node.style.left = x + 'px'; node.style.top = y + 'px'; box.append(node); }
  }
  return box;
}
function renderDocxPict(pict, ctx) {                                                // VML: 그룹, 그림, 직선 지원
  const grp = all(pict, 'group')[0];
  if (grp && !ctx.inCell) return renderVmlGroup(grp, ctx);
  const im = all(pict, 'imagedata')[0];
  if (im) {
    const rel = ctx.rels.get(im.getAttributeNS(NS_R, 'id')), st = cssStyleMap(im.parentElement);
    const url = rel && ctx.urls.get(rel.path);
    return url ? imgEl(url, pt2px(st.width), pt2px(st.height)) : missing('[이미지 없음]');
  }
  const ln = all(pict, 'line')[0];
  if (ln && !ctx.inCell && ln.getAttribute('from') && ln.getAttribute('to')) {
    const [fx, fy] = ln.getAttribute('from').split(',').map(pt2px), [tx, ty] = ln.getAttribute('to').split(',').map(pt2px);
    const wgt = Math.max(1, pt2px(ln.getAttribute('strokeweight') || '0.75pt')), col = hex((ln.getAttribute('strokecolor') || '').slice(1)) || '#000';
    const el = document.createElement('span');
    el.style.cssText = `position:absolute;display:block;left:${fx - ctx.paraLeftPx}px;top:${fy}px;`;
    if (Math.abs(ty - fy) > Math.abs(tx - fx)) { el.style.height = ty - fy + 'px'; el.style.borderLeft = `${wgt}px solid ${col}`; }
    else { el.style.width = tx - fx + 'px'; el.style.borderTop = `${wgt}px solid ${col}`; }
    if (ctx.curBox) ctx.curBox.style.minHeight = Math.max(fy, ty) + wgt + 2 + 'px';
    return el;
  }
  return null;
}
function renderDocxRun(r, ctx, base, out) {
  const rPr = first(r, 'rPr');
  const rp = mergeProps(base, ...styleChain(ctx.styles, wa(first(rPr, 'rStyle'), 'val')).map((c) => c.r), parseRPr(rPr));
  const span = document.createElement('span');
  for (const n of r.children) {
    switch (n.localName) {
      case 't': span.append(n.textContent); break;
      case 'tab': span.append('\t'); break;
      case 'noBreakHyphen': span.append('\u2011'); break;
      case 'br': case 'cr':
        if (wa(n, 'type') === 'page') { const d = document.createElement('div'); d.className = 'docx-pagebreak'; span.append(d); }
        else span.append(document.createElement('br'));
        break;
      case 'sym': { const c = parseInt(wa(n, 'char'), 16); span.append(c >= 0xF000 ? '\u25A1' : String.fromCharCode(c)); break; }
      case 'drawing': { const d = renderDocxDrawing(n, ctx); if (d) span.append(d); break; }
      case 'pict': { const d = renderDocxPict(n, ctx); if (d) span.append(d); break; }
      case 'AlternateContent': {
        const src = first(n, 'Choice') || n, dr = all(src, 'drawing')[0], pc = all(src, 'pict')[0];
        const d = dr ? renderDocxDrawing(dr, ctx) : pc ? renderDocxPict(pc, ctx) : null;
        if (d) span.append(d);
        break;
      }
      case 'footnoteReference': case 'endnoteReference': {
        const kind = n.localName === 'footnoteReference' ? 'footnote' : 'endnote', id = wa(n, 'id');
        const order = ctx.noteOrder[kind]; let i = order.indexOf(id); if (i < 0) i = order.push(id) - 1;
        const sup = document.createElement('sup'); sup.textContent = i + 1; span.append(sup);
        break;
      }
    }
  }
  if (!span.childNodes.length) return;
  applyDocxChar(span, rp, span.textContent);
  out.append(span);
}
function safeHref(h) { return /^(https?:|mailto:)/i.test(h || '') ? h : null; }
function renderInline(parent, ctx, base, out) {                                      // 문단 안의 r / hyperlink / ins / sdt ...
  for (const c of parent.children) {
    const n = c.localName;
    if (n === 'r') renderDocxRun(c, ctx, base, out);
    else if (n === 'hyperlink') {
      const rel = ctx.rels.get(c.getAttributeNS(NS_R, 'id')), href = rel && rel.ext ? safeHref(rel.target) : null;
      const box = href ? document.createElement('a') : document.createElement('span');
      if (href) { box.href = href; box.target = '_blank'; box.rel = 'noopener noreferrer'; }
      renderInline(c, ctx, base, box); out.append(box);
    } else if (n === 'ins' || n === 'smartTag' || n === 'fldSimple' || n === 'moveTo') renderInline(c, ctx, base, out);
    else if (n === 'sdt') { const sc = first(c, 'sdtContent'); if (sc) renderInline(sc, ctx, base, out); }
  }
}
function renderDocxPara(p, ctx) {
  const pPr = first(p, 'pPr');
  const direct = parsePPr(pPr);
  const sid = wa(first(pPr, 'pStyle'), 'val') || ctx.styles.defPara;
  const chain = styleChain(ctx.styles, sid);
  const rBase = mergeProps(ctx.styles.defaults.r, ...chain.map((c) => c.r));
  let pp = mergeProps(ctx.styles.defaults.p, ...chain.map((c) => c.p));
  const nb = nextLabel(ctx, direct.numPr || pp.numPr);
  pp = mergeProps(pp, nb && nb.lv.p, direct);

  const box = document.createElement('div'); box.className = 'docx-p';
  const s = box.style;
  const mark = mergeProps(rBase, parseRPr(first(pPr, 'rPr')));                        // 문단 부호 서식 = 빈 줄 높이
  s.fontSize = (mark.size || 10) + 'pt';
  if (chain.some((c) => /^(heading|제목)\s*\d/.test(c.name))) box.setAttribute('role', 'heading');
  if (pp.jc) s.textAlign = { left: 'left', start: 'left', center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'justify' }[pp.jc] || '';
  if (pp.left) s.marginLeft = tw(pp.left) + 'px';
  if (pp.right) s.marginRight = tw(pp.right) + 'px';
  const first1 = pp.hanging ? -pp.hanging : pp.firstLine || 0;
  if (first1) s.textIndent = tw(first1) + 'px';
  if (pp.before) s.paddingTop = pp.before / 20 + 'pt';                                // margin은 겹쳐지므로 padding 사용
  if (pp.after) s.paddingBottom = pp.after / 20 + 'pt';
  const lineRule = pp.lineRule || 'auto';
  if (pp.line) {
    if (lineRule === 'auto') s.lineHeight = (pp.line / 240) * 1.2;
    else if (lineRule === 'exact') s.lineHeight = pp.line / 20 + 'pt';
    else s.minHeight = pp.line / 20 + 'pt';
  }
  if (pp.fill) s.backgroundColor = pp.fill;
  if (pp.pageBreakBefore) box.classList.add('docx-break');

  ctx.paraLeftPx = ctx.page.left + tw(pp.left || 0);
  ctx.curBox = box;
  if (nb) {
    const lab = document.createElement('span');
    lab.textContent = nb.label; lab.style.display = 'inline-block'; lab.style.minWidth = tw(pp.hanging || 0) + 'px';
    applyDocxChar(lab, mergeProps(rBase, { fontEa: undefined, font: undefined }, nb.lv.r), nb.label);
    box.append(lab);
  }
  renderInline(p, ctx, rBase, box);
  if (!box.childNodes.length) box.append(document.createElement('br'));              // 빈 문단도 줄 높이 유지
  if (s.textAlign === 'justify' && /\t| {3,}/.test(box.textContent)) s.textAlign = 'left'; // 공백·탭으로 칸을 맞춘 줄은 늘리지 않음
  return box;
}
function renderDocxTable(tbl, ctx) {
  const tblPr = first(tbl, 'tblPr');
  const styleB = mergeProps(...styleChain(ctx.styles, wa(first(tblPr, 'tblStyle'), 'val')).map((c) => c.borders));
  const tb = Object.assign({}, styleB, parseBorders(first(tblPr, 'tblBorders')));
  const mar = first(tblPr, 'tblCellMar');
  const cm = (k, d) => { const e = first(mar, k); return e ? tw(wa(e, 'w')) : d; };
  const defPad = { top: cm('top', 0), right: cm('right', 108 / 15), bottom: cm('bottom', 0), left: cm('left', 108 / 15) };
  const grid = kids(first(tbl, 'tblGrid') || tbl, 'gridCol').map((g) => tw(wa(g, 'w')));
  const t = document.createElement('table'); t.className = 'docx-tbl';
  if (grid.length) {
    t.style.tableLayout = 'fixed'; t.style.width = grid.reduce((a, b) => a + b, 0) + 'px';
    const cg = document.createElement('colgroup');
    for (const w of grid) { const c = document.createElement('col'); c.style.width = w + 'px'; cg.append(c); }
    t.append(cg);
  }
  const jc = wa(first(tblPr, 'jc'), 'val'); if (jc === 'center') t.style.margin = '0 auto'; else if (jc === 'right' || jc === 'end') t.style.marginLeft = 'auto';
  const ind = first(tblPr, 'tblInd'); if (ind) t.style.marginLeft = tw(wa(ind, 'w')) + 'px';
  const rows = kids(tbl, 'tr'), owner = [], prevCell = ctx.inCell;
  rows.forEach((tr, ri) => {
    const row = t.insertRow(); let col = 0;
    const trPr = first(tr, 'trPr');
    col += Number(wa(first(trPr, 'gridBefore'), 'val')) || 0;
    const th = first(trPr, 'trHeight'); if (th) row.style.height = tw(wa(th, 'val')) + 'px';
    for (const tc of kids(tr, 'tc')) {
      const tcPr = first(tc, 'tcPr');
      const span = Number(wa(first(tcPr, 'gridSpan'), 'val')) || 1, vm = first(tcPr, 'vMerge');
      if (vm && wa(vm, 'val') !== 'restart') { if (owner[col]) owner[col].rowSpan += 1; col += span; continue; }
      const td = row.insertCell();
      if (span > 1) td.colSpan = span;
      for (let k = 0; k < span; k++) owner[col + k] = td;
      const w = grid.slice(col, col + span).reduce((a, b) => a + b, 0) || tw(wa(first(tcPr, 'tcW'), 'w'));
      if (w) td.style.width = w + 'px';
      const cb = parseBorders(first(tcPr, 'tcBorders'));
      const last = grid.length && col + span >= grid.length;
      td.style.borderTop = cssBorder(cb.top || (ri === 0 ? tb.top : tb.insideH));
      td.style.borderBottom = cssBorder(cb.bottom || (ri === rows.length - 1 ? tb.bottom : tb.insideH));
      td.style.borderLeft = cssBorder(cb.left || (col === 0 ? tb.left : tb.insideV));
      td.style.borderRight = cssBorder(cb.right || (last ? tb.right : tb.insideV));
      const tm = first(tcPr, 'tcMar'), pad = (k, d) => { const e = first(tm, k); return e ? tw(wa(e, 'w')) : d; };
      td.style.padding = [pad('top', defPad.top), pad('right', defPad.right), pad('bottom', defPad.bottom), pad('left', defPad.left)].map((v) => v + 'px').join(' ');
      const fill = first(tcPr, 'shd') && hex(wa(first(tcPr, 'shd'), 'fill')); if (fill) td.style.backgroundColor = fill;
      const va = wa(first(tcPr, 'vAlign'), 'val'); td.style.verticalAlign = va === 'center' ? 'middle' : va === 'bottom' ? 'bottom' : 'top';
      ctx.inCell = true; renderBlocks(tc, td, ctx); ctx.inCell = prevCell;
      if (!td.childNodes.length) td.append(document.createElement('br'));
      col += span;
    }
  });
  return t;
}
function renderBlocks(parent, out, ctx) {                                            // 본문/셀/텍스트상자 공통
  for (const c of parent.children) {
    if (c.localName === 'p') out.append(renderDocxPara(c, ctx));
    else if (c.localName === 'tbl') out.append(renderDocxTable(c, ctx));
    else if (c.localName === 'sdt') { const sc = first(c, 'sdtContent'); if (sc) renderBlocks(sc, out, ctx); }
  }
}
function renderDocxNotes(ctx, sheet) {
  for (const kind of ['footnote', 'endnote']) {
    const order = ctx.noteOrder[kind];
    if (!order.length) continue;
    const box = document.createElement('div'); box.className = 'docx-notes';
    order.forEach((id, i) => {
      const el = ctx.notes[kind].get(id); if (!el) return;
      const row = document.createElement('div');
      renderBlocks(el, row, ctx);
      const sup = document.createElement('sup'); sup.textContent = i + 1;
      if (row.firstElementChild) row.firstElementChild.prepend(sup, ' ');
      box.append(row);
    });
    sheet.append(box);
  }
}
async function renderDocx(zip, viewer) {
  const docXml = await readXml(zip, 'word/document.xml');
  const opt = async (p, fn) => (findEntry(zip, p) ? fn(await readXml(zip, p)) : null);
  const styles = (await opt('word/styles.xml', parseStyles)) || { defaults: { p: {}, r: {} }, byId: new Map(), defPara: null };
  const numbering = (await opt('word/numbering.xml', parseNumbering)) || { abs: new Map(), nums: new Map() };
  const rels = (await opt('word/_rels/document.xml.rels', parseRels)) || new Map();
  const notes = { footnote: await opt('word/footnotes.xml', parseNotes), endnote: await opt('word/endnotes.xml', parseNotes) };
  notes.footnote = notes.footnote || new Map(); notes.endnote = notes.endnote || new Map();
  const urls = await loadDocxImages(zip);
  currentUrls.push(...urls.values());

  const body = all(docXml, 'body')[0];
  if (!body) throw new Error('문서 본문(w:body)을 찾을 수 없습니다');
  const sect = first(body, 'sectPr'), sz = first(sect, 'pgSz'), mg = first(sect, 'pgMar');
  const page = { w: tw(wa(sz, 'w')) || 794, h: tw(wa(sz, 'h')) || 1123, top: tw(wa(mg, 'top') ?? 1440), right: tw(wa(mg, 'right') ?? 1440), bottom: tw(wa(mg, 'bottom') ?? 1440), left: tw(wa(mg, 'left') ?? 1440) };
  const ctx = { styles, numbering, rels, urls, notes, page, counters: new Map(), noteOrder: { footnote: [], endnote: [] }, inCell: false, paraLeftPx: 0 };

  const sheet = document.createElement('section'); sheet.className = 'sheet docx-sheet';
  sheet.style.width = page.w + 'px'; sheet.style.minHeight = page.h + 'px';
  sheet.style.padding = [page.top, page.right, page.bottom, page.left].map((v) => v + 'px').join(' ');
  renderBlocks(body, sheet, ctx);
  renderDocxNotes(ctx, sheet);
  viewer.append(sheet);
  return `문단 ${all(docXml, 'p').length}개, 표 ${all(docXml, 'tbl').length}개, 이미지 ${urls.size}개`;
}

// ---- 5단계: 파일 열기 / UI (HWPX, DOCX 공통) ----
let currentUrls = [];

async function renderHwpx(zip, viewer) {
  const styles = parseHeader(await readXml(zip, 'Contents/header.xml'));
  const manifest = await loadManifest(zip);
  const urls = await loadImages(zip, manifest);
  currentUrls.push(...urls.values());
  const ctx = { styles, urls };
  for (const path of manifest.spine) viewer.append(renderSection(await readXml(zip, path), ctx));
  return `섹션 ${manifest.spine.length}개, 이미지 ${urls.size}개`;
}

// 확장자가 아니라 ZIP 내부 구조로 형식을 판별한다.
function detectKind(zip) {
  if (findEntry(zip, 'word/document.xml')) return 'docx';
  if (findEntry(zip, 'Contents/header.xml') || findEntry(zip, 'Contents/content.hpf')) return 'hwpx';
  return null;
}

async function openFile(file) {
  const status = $('status'), viewer = $('viewer');
  status.className = ''; status.textContent = `${file.name} 여는 중...`;
  currentUrls.forEach(URL.revokeObjectURL); currentUrls = []; viewer.replaceChildren();
  try {
    if (/\.hwp$/i.test(file.name)) throw new Error('.hwp(바이너리) 형식은 지원하지 않습니다. .hwpx로 저장해 주세요.');
    if (/\.doc$/i.test(file.name)) throw new Error('.doc(바이너리) 형식은 지원하지 않습니다. .docx로 저장해 주세요.');
    let zip;
    try { zip = await JSZip.loadAsync(file); } catch { throw new Error('올바른 문서 파일이 아닙니다. (.hwpx, .docx만 열 수 있습니다)'); }
    const kind = detectKind(zip);
    if (!kind) throw new Error('지원하지 않는 형식입니다. (.hwpx, .docx만 열 수 있습니다)');
    const ext = file.name.split('.').pop().toLowerCase();
    const summary = kind === 'docx' ? await renderDocx(zip, viewer) : await renderHwpx(zip, viewer);
    $('tree').textContent = Object.values(zip.files).map((f) => (f.dir ? f.name : `${f.name}  (${f._data?.uncompressedSize ?? '?'} B)`)).join('\n');
    $('structure').hidden = false;
    const note = ext !== kind ? ` [확장자는 .${ext}지만 내용은 ${kind}입니다]` : '';
    status.textContent = `${file.name} (${kind.toUpperCase()}) - ${summary}${note}`;
  } catch (e) {
    status.className = 'error';
    status.textContent = `열 수 없습니다: ${e.message}`; // 암호화/배포용 문서는 header.xml 접근이 막혀 있을 수 있다
  }
}

$('file').addEventListener('change', (e) => e.target.files[0] && openFile(e.target.files[0]));
const drop = $('drop');
['dragenter', 'dragover'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add('over'); }));
['dragleave', 'drop'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove('over'); }));
drop.addEventListener('drop', (e) => e.dataTransfer.files[0] && openFile(e.dataTransfer.files[0]));
