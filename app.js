'use strict';
/* PPT 备注打印排版 — 纯前端解析 .pptx，渲染幻灯片 + 抽取备注，排版成 A4 打印页 */
(function () {

// ---------- 常量与单位换算 ----------
const EMU_PER_PX = 9525;                 // 914400 EMU/英寸 ÷ 96 px/英寸
const MM = 96 / 25.4;                    // 1mm -> px (96dpi)
const SHEET_PAD_MM = 4;
const emu = v => (parseInt(v, 10) || 0) / EMU_PER_PX;
const ptToPx = sz => (parseFloat(sz) || 1800) / 75;   // sz 为「百分之一磅」
const num = v => (v == null ? 0 : parseInt(v, 10) || 0);

const MIME = { png:'image/png', jpg:'image/jpeg', jpeg:'image/jpeg', gif:'image/gif',
  bmp:'image/bmp', webp:'image/webp', svg:'image/svg+xml', tif:'image/tiff', tiff:'image/tiff' };
const PRESET_COLORS = { black:'000000', white:'FFFFFF', red:'FF0000', green:'008000',
  blue:'0000FF', yellow:'FFFF00', gray:'808080', grey:'808080', orange:'FFA500',
  purple:'800080', darkGray:'A9A9A9', lightGray:'D3D3D3' };

const CJK_FALLBACK = '"PingFang SC","Microsoft YaHei","Hiragino Sans GB","Heiti SC",sans-serif';

// ---------- DOM 速记 ----------
const $ = id => document.getElementById(id);
const dropzone = $('dropzone'), fileInput = $('fileInput'), workspace = $('workspace'),
      preview = $('preview'), statusEl = $('status'), loading = $('loading'), meta = $('meta');

let deck = null;     // { renderW, renderH, pdfW, pdfH, slides:[{domEl, pdfImg, notes, title, index}] }
const opts = { picSource:'render', layout:'side', perPage:3, noteSize:'m', showNum:true, hideEmpty:false };

if(typeof pdfjsLib !== 'undefined')
  pdfjsLib.GlobalWorkerOptions.workerSrc = 'lib/pdf.worker.min.js';

// ---------- XML 工具 ----------
function parseXML(str){ return new DOMParser().parseFromString(str, 'application/xml'); }
function dchildren(el, name){
  const o = []; if(!el) return o;
  for(let c = el.firstElementChild; c; c = c.nextElementSibling) if(c.tagName === name) o.push(c);
  return o;
}
function dchild(el, name){
  if(!el) return null;
  for(let c = el.firstElementChild; c; c = c.nextElementSibling) if(c.tagName === name) return c;
  return null;
}
function deep(el, name){ return el ? Array.from(el.getElementsByTagName(name)) : []; }
function deep1(el, name){ const r = el ? el.getElementsByTagName(name) : null; return r && r[0] || null; }

// ---------- 路径处理 ----------
function resolvePath(baseDir, target){
  if(!target) return '';
  if(target.charAt(0) === '/') return target.slice(1);
  const parts = baseDir.split('/').filter(Boolean);
  target.split('/').forEach(seg => {
    if(seg === '..') parts.pop();
    else if(seg && seg !== '.') parts.push(seg);
  });
  return parts.join('/');
}
function dirOf(p){ const i = p.lastIndexOf('/'); return i < 0 ? '' : p.slice(0, i); }

// ---------- 颜色 ----------
function hexToRgb(h){
  h = (h || '000000').replace('#','');
  if(h.length === 3) h = h[0]+h[0]+h[1]+h[1]+h[2]+h[2];
  return { r:parseInt(h.slice(0,2),16)||0, g:parseInt(h.slice(2,4),16)||0,
           b:parseInt(h.slice(4,6),16)||0, a:1 };
}
const clamp = v => Math.max(0, Math.min(255, Math.round(v)));
function rgbCss(c){
  return c.a < 1 ? `rgba(${clamp(c.r)},${clamp(c.g)},${clamp(c.b)},${c.a.toFixed(3)})`
                 : `rgb(${clamp(c.r)},${clamp(c.g)},${clamp(c.b)})`;
}
function applyMod(c, tag, v){
  if(tag === 'a:tint')   return { r:c.r*v+255*(1-v), g:c.g*v+255*(1-v), b:c.b*v+255*(1-v), a:c.a };
  if(tag === 'a:shade')  return { r:c.r*v, g:c.g*v, b:c.b*v, a:c.a };
  if(tag === 'a:lumMod') return { r:c.r*v, g:c.g*v, b:c.b*v, a:c.a };
  if(tag === 'a:lumOff') return { r:c.r+255*v, g:c.g+255*v, b:c.b+255*v, a:c.a };
  if(tag === 'a:alpha')  return { r:c.r, g:c.g, b:c.b, a:v };
  return c;
}
function schemeHex(name, ctx){
  if(!name) return '000000';
  if(name === 'phClr') return ctx.phClr || '4F81BD';
  const slot = (ctx.clrMap && ctx.clrMap[name]) || name;
  const t = ctx.theme.colors;
  return t[slot] || t[name] || '000000';
}
function colorFromEl(cEl, ctx){
  if(!cEl) return null;
  let hex = null;
  switch(cEl.tagName){
    case 'a:srgbClr':   hex = cEl.getAttribute('val'); break;
    case 'a:sysClr':    hex = cEl.getAttribute('lastClr') || '000000'; break;
    case 'a:schemeClr': hex = schemeHex(cEl.getAttribute('val'), ctx); break;
    case 'a:prstClr':   hex = PRESET_COLORS[cEl.getAttribute('val')] || '000000'; break;
    case 'a:scrgbClr': {
      const pr = a => Math.round((parseFloat(cEl.getAttribute(a))||0)/100000*255);
      hex = [pr('r'),pr('g'),pr('b')].map(n => n.toString(16).padStart(2,'0')).join(''); break;
    }
    default: return null;
  }
  if(hex == null) return null;
  let rgb = hexToRgb(hex);
  for(let m = cEl.firstElementChild; m; m = m.nextElementSibling){
    const raw = m.getAttribute('val');
    if(raw == null) continue;
    rgb = applyMod(rgb, m.tagName, parseFloat(raw)/100000);
  }
  return rgbCss(rgb);
}
function fillFromEl(el, ctx){
  if(!el) return null;
  if(el.tagName === 'a:noFill') return { type:'none' };
  if(el.tagName === 'a:solidFill'){
    const c = colorFromEl(el.firstElementChild, ctx);
    return c ? { type:'solid', css:c } : null;
  }
  if(el.tagName === 'a:gradFill'){
    const stops = [];
    deep(el, 'a:gs').forEach(g => {
      const c = colorFromEl(g.firstElementChild, ctx);
      if(c) stops.push(c + ' ' + ((parseInt(g.getAttribute('pos')||'0',10))/1000) + '%');
    });
    if(!stops.length) return null;
    let ang = 90;
    const lin = deep1(el, 'a:lin');
    if(lin) ang = (parseInt(lin.getAttribute('ang')||'0',10)/60000) + 90;
    return { type:'grad', css:`linear-gradient(${ang}deg,${stops.join(',')})` };
  }
  return null;
}

// ---------- 几何 ----------
function readXfrm(xfrm){
  if(!xfrm) return null;
  const off = dchild(xfrm,'a:off'), ext = dchild(xfrm,'a:ext');
  if(!off || !ext) return null;
  return {
    x: emu(off.getAttribute('x')),  y: emu(off.getAttribute('y')),
    w: emu(ext.getAttribute('cx')), h: emu(ext.getAttribute('cy')),
    rot: xfrm.getAttribute('rot') ? parseInt(xfrm.getAttribute('rot'),10)/60000 : 0,
    flipH: xfrm.getAttribute('flipH') === '1',
    flipV: xfrm.getAttribute('flipV') === '1'
  };
}
function absStyle(g){
  let s = `position:absolute;left:${g.x}px;top:${g.y}px;width:${g.w}px;height:${g.h}px;`;
  if(g.rot) s += `transform:rotate(${g.rot}deg);`;
  return s;
}

// ---------- 占位符 ----------
function phFamily(t){ return (t === 'title' || t === 'ctrTitle') ? 'title' : 'body'; }
function readPh(sp){
  const nv = dchild(sp,'p:nvSpPr') || dchild(sp,'p:nvPicPr') || dchild(sp,'p:nvCxnSpPr')
           || dchild(sp,'p:nvGraphicFramePr');
  const ph = nv ? deep1(dchild(nv,'p:nvPr'), 'p:ph') : null;
  if(!ph) return null;
  return { type: ph.getAttribute('type') || 'body', idx: ph.getAttribute('idx') || '' };
}
function collectPlaceholders(xmlDoc){
  const list = [];
  const tree = deep1(xmlDoc, 'p:spTree');
  if(!tree) return list;
  dchildren(tree, 'p:sp').forEach(sp => {
    const ph = readPh(sp);
    if(!ph) return;
    const spPr = dchild(sp,'p:spPr');
    const txBody = dchild(sp,'p:txBody');
    list.push({
      type: ph.type, idx: ph.idx,
      geo: readXfrm(dchild(spPr,'a:xfrm')),
      bodyPr: dchild(txBody,'a:bodyPr'),
      lstStyle: parseLstStyle(dchild(txBody,'a:lstStyle'))
    });
  });
  return list;
}
function matchPh(list, type, idx){
  if(!list) return null;
  const fam = phFamily(type);
  return list.find(p => p.idx === idx && phFamily(p.type) === fam)
      || list.find(p => p.type === type)
      || list.find(p => phFamily(p.type) === fam)
      || list.find(p => p.idx === idx && idx !== '')
      || null;
}

// ---------- 文本样式 ----------
function parseRunProps(rPr){
  const o = {};
  if(!rPr) return o;
  if(rPr.getAttribute('sz') != null) o.sz = parseInt(rPr.getAttribute('sz'),10);
  if(rPr.getAttribute('b') != null) o.bold = rPr.getAttribute('b') === '1';
  if(rPr.getAttribute('i') != null) o.italic = rPr.getAttribute('i') === '1';
  const u = rPr.getAttribute('u'); if(u && u !== 'none') o.underline = true;
  const sf = dchild(rPr,'a:solidFill');
  if(sf && sf.firstElementChild) o.colorEl = sf.firstElementChild;
  const latin = dchild(rPr,'a:latin');
  if(latin && latin.getAttribute('typeface')) o.font = latin.getAttribute('typeface');
  return o;
}
function parseLvlPr(el){
  const o = {};
  if(!el) return o;
  if(el.getAttribute('algn'))   o.algn = el.getAttribute('algn');
  if(el.getAttribute('marL') != null)   o.marL = el.getAttribute('marL');
  if(el.getAttribute('indent') != null) o.indent = el.getAttribute('indent');
  if(dchild(el,'a:buNone')) o.buNone = true;
  const bc = dchild(el,'a:buChar'); if(bc) o.buChar = bc.getAttribute('char');
  if(dchild(el,'a:buAutoNum')) o.buAutoNum = true;
  Object.assign(o, parseRunProps(dchild(el,'a:defRPr')));
  return o;
}
function parseLstStyle(lst){
  const o = {};
  if(!lst) return o;
  for(let c = lst.firstElementChild; c; c = c.nextElementSibling){
    const m = c.tagName.match(/lvl(\d)pPr/);
    if(m) o[parseInt(m[1],10)-1] = parseLvlPr(c);
  }
  return o;
}
function parseTxStyles(masterDoc){
  const out = { title:{}, body:{}, other:{} };
  const ts = deep1(masterDoc, 'p:txStyles');
  if(!ts) return out;
  [['p:titleStyle','title'],['p:bodyStyle','body'],['p:otherStyle','other']].forEach(([tag,key]) => {
    const sty = dchild(ts, tag);
    if(!sty) return;
    for(let c = sty.firstElementChild; c; c = c.nextElementSibling){
      const m = c.tagName.match(/lvl(\d)pPr/);
      if(m) out[key][parseInt(m[1],10)-1] = parseLvlPr(c);
    }
  });
  return out;
}

// ---------- 包(zip)读取 ----------
async function readText(zip, path){
  const f = zip.file(path);
  return f ? f.async('string') : null;
}
async function readRels(zip, partPath){
  const base = dirOf(partPath), nameOnly = partPath.slice(base.length ? base.length+1 : 0);
  const relsPath = (base ? base + '/' : '') + '_rels/' + nameOnly + '.rels';
  const txt = await readText(zip, relsPath);
  const map = {};
  if(!txt) return map;
  deep(parseXML(txt), 'Relationship').forEach(r => {
    map[r.getAttribute('Id')] = { type:r.getAttribute('Type'), target:r.getAttribute('Target') };
  });
  return map;
}
function relByType(rels, kind){
  for(const id in rels) if(rels[id].type && rels[id].type.indexOf('/'+kind) >= 0) return rels[id];
  return null;
}

// ============================================================
//  主流程：加载 pptx
// ============================================================
async function loadPptx(buffer, fileName){
  if(typeof JSZip === 'undefined') throw new Error('依赖库未能加载，请确认 lib/jszip.min.js 存在。');
  let zip;
  try { zip = await JSZip.loadAsync(buffer); }
  catch(e){ throw new Error('无法读取该文件。请确认它是有效的 .pptx 文件（旧版 .ppt 不支持，请先另存为 .pptx）。'); }

  const presXmlText = await readText(zip, 'ppt/presentation.xml');
  if(!presXmlText) throw new Error('这不是一个有效的 PowerPoint 文件（缺少 presentation.xml）。');
  const presDoc = parseXML(presXmlText);

  const sldSz = deep1(presDoc, 'p:sldSz');
  const slideW = sldSz ? emu(sldSz.getAttribute('cx')) : 960;
  const slideH = sldSz ? emu(sldSz.getAttribute('cy')) : 720;

  const presRels = await readRels(zip, 'ppt/presentation.xml');

  // 预加载所有图片为 dataURL
  const images = {};
  await Promise.all(Object.keys(zip.files)
    .filter(p => p.indexOf('ppt/media/') === 0 && !zip.files[p].dir)
    .map(async p => {
      const ext = p.split('.').pop().toLowerCase();
      if(!MIME[ext]) return;
      images[p] = 'data:' + MIME[ext] + ';base64,' + await zip.file(p).async('base64');
    }));

  // 缓存 layout / master / theme
  const cacheLayout = {}, cacheMaster = {}, cacheTheme = {};

  async function getTheme(path){
    if(cacheTheme[path]) return cacheTheme[path];
    const txt = await readText(zip, path);
    const colors = {}; let majorFont = 'Calibri', minorFont = 'Calibri';
    if(txt){
      const doc = parseXML(txt);
      const cs = deep1(doc,'a:clrScheme');
      if(cs) for(let c = cs.firstElementChild; c; c = c.nextElementSibling){
        const slot = c.tagName.replace('a:','');
        const inner = c.firstElementChild;
        if(!inner) continue;
        colors[slot] = inner.tagName === 'a:sysClr'
          ? (inner.getAttribute('lastClr') || '000000')
          : (inner.getAttribute('val') || '000000');
      }
      const mj = deep1(doc,'a:majorFont'), mn = deep1(doc,'a:minorFont');
      if(mj && dchild(mj,'a:latin')) majorFont = dchild(mj,'a:latin').getAttribute('typeface') || majorFont;
      if(mn && dchild(mn,'a:latin')) minorFont = dchild(mn,'a:latin').getAttribute('typeface') || minorFont;
    }
    return (cacheTheme[path] = { colors, majorFont, minorFont });
  }

  async function getMaster(path){
    if(cacheMaster[path]) return cacheMaster[path];
    const doc = parseXML(await readText(zip, path));
    const rels = await readRels(zip, path);
    const themeRel = relByType(rels,'theme');
    const theme = await getTheme(themeRel ? resolvePath(dirOf(path), themeRel.target) : '');
    const clrMapEl = deep1(doc,'p:clrMap');
    const clrMap = {};
    if(clrMapEl) for(const a of clrMapEl.attributes) clrMap[a.name] = a.value;
    const rec = {
      doc, clrMap, theme,
      rels, dir: dirOf(path),
      placeholders: collectPlaceholders(doc),
      txStyles: parseTxStyles(doc),
      bg: dchild(deep1(doc,'p:cSld'),'p:bg')
    };
    return (cacheMaster[path] = rec);
  }

  async function getLayout(path){
    if(cacheLayout[path]) return cacheLayout[path];
    const doc = parseXML(await readText(zip, path));
    const rels = await readRels(zip, path);
    const masterRel = relByType(rels,'slideMaster');
    const master = await getMaster(resolvePath(dirOf(path), masterRel.target));
    const rec = {
      doc, master,
      rels, dir: dirOf(path),
      placeholders: collectPlaceholders(doc),
      bg: dchild(deep1(doc,'p:cSld'),'p:bg')
    };
    return (cacheLayout[path] = rec);
  }

  // 逐张幻灯片
  const sldIds = deep(presDoc, 'p:sldId');
  const slides = [];
  for(let i = 0; i < sldIds.length; i++){
    const rid = sldIds[i].getAttribute('r:id');
    const rel = presRels[rid];
    if(!rel) continue;
    const slidePath = resolvePath('ppt', rel.target);
    try {
      const slideDoc = parseXML(await readText(zip, slidePath));
      const slideRels = await readRels(zip, slidePath);
      const layoutRel = relByType(slideRels,'slideLayout');
      const layout = await getLayout(resolvePath(dirOf(slidePath), layoutRel.target));

      const ctx = {
        theme: layout.master.theme,
        clrMap: layout.master.clrMap,
        master: layout.master,
        layout: layout,
        slideDir: dirOf(slidePath),
        slideRels, images, slideW, slideH
      };

      const nativeEl = renderSlide(slideDoc, ctx);

      // 备注
      let notes = [];
      const notesRel = relByType(slideRels,'notesSlide');
      if(notesRel){
        const notesTxt = await readText(zip, resolvePath(dirOf(slidePath), notesRel.target));
        if(notesTxt) notes = extractNotes(parseXML(notesTxt));
      }
      slides.push({ domEl: nativeEl, pdfImg: null, notes,
        title: slideTitle(slideDoc), index: slides.length+1 });
    } catch(err){
      console.error('幻灯片解析失败 ' + slidePath, err);
      const ph = document.createElement('div');
      ph.className = 'slide-native';
      ph.style.cssText = `width:${slideW}px;height:${slideH}px;display:flex;align-items:center;`
        + `justify-content:center;color:#b3bac8;font-size:28px;`;
      ph.textContent = '（此页无法渲染）';
      slides.push({ domEl: ph, pdfImg: null, notes: [], title:'', index: slides.length+1 });
    }
  }
  if(!slides.length) throw new Error('文件里没有找到任何幻灯片。');
  return { renderW: slideW, renderH: slideH, pdfW: 0, pdfH: 0, slides };
}

// ============================================================
//  渲染：幻灯片
// ============================================================
function renderSlide(slideDoc, ctx){
  const root = document.createElement('div');
  root.className = 'slide-native';
  root.style.width = ctx.slideW + 'px';
  root.style.height = ctx.slideH + 'px';
  root.style.background = resolveBg(slideDoc, ctx);

  // 背景层：母版 + 版式中的「非占位符」装饰图形（logo、色块、固定文字等）
  const showMaster = (deep1(slideDoc,'p:sld') || {}).getAttribute
    ? deep1(slideDoc,'p:sld').getAttribute('showMasterSp') !== '0' : true;
  if(showMaster){
    safeRenderTree(deep1(ctx.master.doc,'p:spTree'), ctx, root, true,
      { rels: ctx.master.rels, dir: ctx.master.dir });
    safeRenderTree(deep1(ctx.layout.doc,'p:spTree'), ctx, root, true,
      { rels: ctx.layout.rels, dir: ctx.layout.dir });
  }
  // 幻灯片自身的图形
  safeRenderTree(deep1(slideDoc,'p:spTree'), ctx, root, false,
    { rels: ctx.slideRels, dir: ctx.slideDir });
  return root;
}
function resolveBg(slideDoc, ctx){
  const pick = bg => {
    if(!bg) return null;
    const pr = dchild(bg,'p:bgPr');
    if(pr){ const f = fillFromEl(pr.firstElementChild, ctx); if(f && f.css) return f.css; }
    const ref = dchild(bg,'p:bgRef');
    if(ref){ const c = colorFromEl(ref.firstElementChild, ctx); if(c) return c; }
    return null;
  };
  return pick(dchild(deep1(slideDoc,'p:cSld'),'p:bg'))
      || pick(ctx.layout.bg) || pick(ctx.master.bg) || '#ffffff';
}
function safeRenderTree(tree, ctx, parent, skipPh, src){
  if(!tree) return;
  for(let el = tree.firstElementChild; el; el = el.nextElementSibling){
    try {
      let node = null;
      if(el.tagName === 'p:sp'){
        if(skipPh && readPh(el)) continue;
        node = renderSp(el, ctx);
      } else if(el.tagName === 'p:pic'){
        node = renderPic(el, ctx, src);
      } else if(el.tagName === 'p:cxnSp'){
        node = renderSp(el, ctx);
      } else if(el.tagName === 'p:graphicFrame'){
        node = renderGraphicFrame(el, ctx);
      } else if(el.tagName === 'p:grpSp'){
        node = renderGrp(el, ctx, skipPh, src);
      }
      if(node) parent.appendChild(node);
    } catch(e){ console.warn('图形渲染跳过', e); }
  }
}

function renderGrp(grp, ctx, skipPh, src){
  const xfrm = dchild(dchild(grp,'p:grpSpPr'),'a:xfrm');
  if(!xfrm) return null;
  const off = dchild(xfrm,'a:off'), ext = dchild(xfrm,'a:ext'),
        chOff = dchild(xfrm,'a:chOff'), chExt = dchild(xfrm,'a:chExt');
  if(!off || !ext) return null;
  const x = emu(off.getAttribute('x')), y = emu(off.getAttribute('y'));
  const w = emu(ext.getAttribute('cx')), h = emu(ext.getAttribute('cy'));
  const cw = chExt ? emu(chExt.getAttribute('cx')) || w : w;
  const ch = chExt ? emu(chExt.getAttribute('cy')) || h : h;
  const cx0 = chOff ? emu(chOff.getAttribute('x')) : 0;
  const cy0 = chOff ? emu(chOff.getAttribute('y')) : 0;
  const box = document.createElement('div');
  box.style.cssText = `position:absolute;left:${x}px;top:${y}px;width:${w}px;height:${h}px;`;
  const inner = document.createElement('div');
  const sx = cw ? w/cw : 1, sy = ch ? h/ch : 1;
  inner.style.cssText = `position:absolute;left:0;top:0;width:${cw}px;height:${ch}px;`
    + `transform:translate(${-cx0*sx}px,${-cy0*sy}px) scale(${sx},${sy});transform-origin:top left;`;
  safeRenderTree(grp, ctx, inner, skipPh, src);
  box.appendChild(inner);
  return box;
}

function renderSp(sp, ctx){
  const ph = readPh(sp);
  const spPr = dchild(sp,'p:spPr');
  let geo = readXfrm(dchild(spPr,'a:xfrm'));
  let layoutPh = null, masterPh = null;
  if(ph){
    layoutPh = matchPh(ctx.layout.placeholders, ph.type, ph.idx);
    masterPh = matchPh(ctx.master.placeholders, ph.type, ph.idx);
    if(!geo) geo = (layoutPh && layoutPh.geo) || (masterPh && masterPh.geo) || null;
  }
  if(!geo) return null;

  const box = document.createElement('div');
  box.style.cssText = absStyle(geo);

  // 填充
  const styleEl = dchild(sp,'p:style');
  let fill = fillFromEl(firstFillChild(spPr), ctx);
  if(!fill && styleEl){
    const fr = dchild(styleEl,'a:fillRef');
    if(fr && fr.getAttribute('idx') !== '0'){
      const c = colorFromEl(fr.firstElementChild, ctx);
      if(c) fill = { type:'solid', css:c };
    }
  }
  if(fill && fill.type !== 'none' && fill.css) box.style.background = fill.css;

  // 描边
  const ln = dchild(spPr,'a:ln');
  if(ln){
    const lf = fillFromEl(firstFillChild(ln), ctx);
    if(lf && lf.type === 'solid'){
      const wpx = ln.getAttribute('w') ? Math.max(1, emu(ln.getAttribute('w'))) : 1;
      box.style.border = `${wpx}px solid ${lf.css}`;
    }
  }

  // 形状外观
  const prst = (deep1(spPr,'a:prstGeom') || {}).getAttribute
    ? deep1(spPr,'a:prstGeom').getAttribute('prst') : null;
  if(prst === 'roundRect') box.style.borderRadius = Math.min(geo.w,geo.h)*0.14 + 'px';
  else if(prst === 'ellipse' || prst === 'oval') box.style.borderRadius = '50%';

  // 文本
  const txBody = dchild(sp,'p:txBody');
  if(txBody){
    const fontRefColor = styleEl ? colorFromEl((dchild(styleEl,'a:fontRef')||{}).firstElementChild, ctx) : null;
    box.appendChild(renderTextBody(txBody, ctx, {
      ph, layoutPh, masterPh, fontRefColor
    }));
  }
  return box;
}
function firstFillChild(el){
  if(!el) return null;
  for(let c = el.firstElementChild; c; c = c.nextElementSibling)
    if(/^a:(solid|grad|no|blip|patt)Fill$/.test(c.tagName)) return c;
  return null;
}

function renderPic(pic, ctx, src){
  const spPr = dchild(pic,'p:spPr');
  const geo = readXfrm(dchild(spPr,'a:xfrm'));
  if(!geo) return null;
  const box = document.createElement('div');
  box.style.cssText = absStyle(geo) + 'overflow:hidden;';
  const blip = deep1(pic,'a:blip');
  if(blip){
    const embed = blip.getAttribute('r:embed') || blip.getAttribute('r:link');
    const rels = (src && src.rels) || ctx.slideRels;
    const dir = (src && src.dir) || ctx.slideDir;
    const rel = rels[embed];
    if(rel){
      const imgSrc = ctx.images[resolvePath(dir, rel.target)];
      if(imgSrc){
        const img = document.createElement('img');
        img.src = imgSrc;
        img.style.cssText = 'width:100%;height:100%;object-fit:fill;display:block;';
        if(geo.flipH || geo.flipV)
          img.style.transform = `scale(${geo.flipH?-1:1},${geo.flipV?-1:1})`;
        box.appendChild(img);
      }
    }
  }
  const prst = deep1(spPr,'a:prstGeom');
  if(prst && (prst.getAttribute('prst')==='ellipse')) box.style.borderRadius = '50%';
  else if(prst && prst.getAttribute('prst')==='roundRect')
    box.style.borderRadius = Math.min(geo.w,geo.h)*0.12 + 'px';
  return box;
}

function renderGraphicFrame(gf, ctx){
  const xfrm = dchild(gf,'p:xfrm');
  const geo = readXfrm(xfrm);
  if(!geo) return null;
  const box = document.createElement('div');
  box.style.cssText = absStyle(geo);
  const tbl = deep1(gf,'a:tbl');
  if(tbl){ try { box.appendChild(renderTable(tbl, ctx, geo)); return box; } catch(e){} }
  // 图表等：占位提示
  box.style.cssText += 'border:1px dashed #b9c2d0;display:flex;align-items:center;'
    + 'justify-content:center;color:#9aa3b4;';
  const label = document.createElement('div');
  label.textContent = deep1(gf,'c:chart') ? '[ 图表 ]' : '[ 内嵌对象 ]';
  label.style.fontSize = Math.min(28, geo.h/4) + 'px';
  box.appendChild(label);
  return box;
}
function renderTable(tbl, ctx, geo){
  const cols = deep(tbl,'a:gridCol').map(c => emu(c.getAttribute('w')));
  const total = cols.reduce((a,b)=>a+b,0) || geo.w;
  const t = document.createElement('table');
  t.style.cssText = `border-collapse:collapse;width:${geo.w}px;height:${geo.h}px;table-layout:fixed;`;
  dchildren(tbl,'a:tr').forEach(tr => {
    const row = document.createElement('tr');
    const rh = emu(tr.getAttribute('h'));
    if(rh) row.style.height = rh + 'px';
    dchildren(tr,'a:tc').forEach((tc, ci) => {
      const td = document.createElement('td');
      td.style.cssText = 'border:1px solid #b9c2d0;padding:3px 6px;vertical-align:middle;'
        + 'overflow:hidden;font-size:14px;';
      if(cols[ci]) td.style.width = (cols[ci]/total*100) + '%';
      const tcPr = dchild(tc,'a:tcPr');
      const f = fillFromEl(firstFillChild(tcPr), ctx);
      if(f && f.css) td.style.background = f.css;
      let txt = '';
      deep(tc,'a:t').forEach(n => txt += n.textContent);
      td.textContent = txt;
      row.appendChild(td);
    });
    t.appendChild(row);
  });
  return t;
}

// ---------- 文本主体 ----------
function renderTextBody(txBody, ctx, info){
  const bodyPr = dchild(txBody,'a:bodyPr');
  const anchorOf = el => el && el.getAttribute && el.getAttribute('anchor');
  const anchor = anchorOf(bodyPr)
    || anchorOf(info.layoutPh && info.layoutPh.bodyPr)
    || anchorOf(info.masterPh && info.masterPh.bodyPr) || 't';
  const insAttr = (el, a, def) =>
    (el && el.getAttribute(a) != null) ? emu(el.getAttribute(a)) : def;
  const ins = {
    l: insAttr(bodyPr,'lIns', 9.6), r: insAttr(bodyPr,'rIns', 9.6),
    t: insAttr(bodyPr,'tIns', 4.8), b: insAttr(bodyPr,'bIns', 4.8)
  };
  const wrap = document.createElement('div');
  wrap.style.cssText = 'position:absolute;left:0;top:0;width:100%;height:100%;'
    + 'display:flex;flex-direction:column;overflow:hidden;'
    + 'justify-content:' + (anchor==='ctr'?'center':anchor==='b'?'flex-end':'flex-start') + ';'
    + `padding:${ins.t}px ${ins.r}px ${ins.b}px ${ins.l}px;`;

  const phType = info.ph ? info.ph.type : null;
  const phClass = !info.ph ? 'other' : (phFamily(phType)==='title' ? 'title' : 'body');
  const shapeLst = parseLstStyle(dchild(txBody,'a:lstStyle'));

  dchildren(txBody,'a:p').forEach((p, i) => {
    wrap.appendChild(renderPara(p, ctx, { phClass, info, shapeLst }, i));
  });
  return wrap;
}
function effectiveLevel(ctx, env, lvl){
  const s = {};
  const tx = ctx.master.txStyles[env.phClass] || {};
  Object.assign(s, tx[lvl] || tx[0] || {});
  const merge = lst => { if(lst && lst[lvl]) Object.assign(s, lst[lvl]); };
  merge(env.info.masterPh && env.info.masterPh.lstStyle);
  merge(env.info.layoutPh && env.info.layoutPh.lstStyle);
  merge(env.shapeLst);
  return s;
}
const ALIGN_CSS = { l:'left', ctr:'center', r:'right', just:'justify', dist:'justify' };

function renderPara(p, ctx, env, index){
  const pPr = dchild(p,'a:pPr');
  const lvl = pPr && pPr.getAttribute('lvl') ? parseInt(pPr.getAttribute('lvl'),10) : 0;
  const base = Object.assign({}, effectiveLevel(ctx, env, lvl), parseLvlPr(pPr));

  const div = document.createElement('div');
  const algn = base.algn || (env.phClass==='title' ? 'ctr' : 'l');
  div.style.textAlign = ALIGN_CSS[algn] || 'left';
  div.style.lineHeight = '1.22';
  if(index > 0) div.style.marginTop = '0.16em';
  const marL = emu(base.marL || 0);
  const indent = emu(base.indent || 0);
  if(marL) div.style.paddingLeft = marL + 'px';

  const wantBullet = !base.buNone && env.phClass === 'body';
  if(wantBullet){
    const b = document.createElement('span');
    b.textContent = base.buChar || '•';
    b.style.cssText = `display:inline-block;width:${Math.max(0,-indent)}px;`
      + `margin-left:${indent}px;`;
    if(base.sz) b.style.fontSize = ptToPx(base.sz) + 'px';
    div.appendChild(b);
  }

  let textCount = 0;
  for(let node = p.firstElementChild; node; node = node.nextElementSibling){
    if(node.tagName === 'a:br'){ div.appendChild(document.createElement('br')); continue; }
    if(node.tagName !== 'a:r' && node.tagName !== 'a:fld') continue;
    const tEl = dchild(node,'a:t');
    const text = tEl ? tEl.textContent : '';
    if(!text) continue;
    const rp = Object.assign({}, base, parseRunProps(dchild(node,'a:rPr')));
    const span = document.createElement('span');
    span.textContent = text;
    span.style.fontSize = ptToPx(rp.sz || 1800) + 'px';
    if(rp.bold) span.style.fontWeight = '700';
    if(rp.italic) span.style.fontStyle = 'italic';
    if(rp.underline) span.style.textDecoration = 'underline';
    let color = colorFromEl(rp.colorEl, ctx);
    if(!color) color = (env.phClass==='other' && env.info.fontRefColor) || '#000000';
    span.style.color = color;
    span.style.fontFamily = '"' + (rp.font || ctx.theme.minorFont || 'Calibri') + '",' + CJK_FALLBACK;
    div.appendChild(span);
    textCount++;
  }
  if(!textCount && !wantBullet){
    div.innerHTML = '&#8203;';
    div.style.minHeight = ptToPx(base.sz || 1800) + 'px';
  }
  return div;
}

// ---------- 备注 / 标题抽取 ----------
function extractNotes(notesDoc){
  const sps = deep(notesDoc,'p:sp');
  let bodySp = null;
  for(const sp of sps){
    const ph = deep1(sp,'p:ph');
    if(ph && ph.getAttribute('type') === 'body'){ bodySp = sp; break; }
  }
  if(!bodySp){
    for(const sp of sps){
      const ph = deep1(sp,'p:ph');
      const t = ph ? ph.getAttribute('type') : null;
      if(t === 'sldNum' || t === 'sldImg' || t === 'dt' || t === 'ftr') continue;
      if(deep(sp,'a:t').length){ bodySp = sp; break; }
    }
  }
  if(!bodySp) return [];
  const paras = [];
  deep(dchild(bodySp,'p:txBody'),'a:p').forEach(p => {
    let s = '';
    for(let n = p.firstElementChild; n; n = n.nextElementSibling){
      if(n.tagName === 'a:r' || n.tagName === 'a:fld'){
        const t = dchild(n,'a:t'); if(t) s += t.textContent;
      } else if(n.tagName === 'a:br') s += '\n';
    }
    paras.push(s);
  });
  while(paras.length && !paras[paras.length-1].trim()) paras.pop();
  while(paras.length && !paras[0].trim()) paras.shift();
  return paras;
}
function slideTitle(slideDoc){
  for(const sp of deep(slideDoc,'p:sp')){
    const ph = deep1(sp,'p:ph');
    if(ph && (ph.getAttribute('type')==='title' || ph.getAttribute('type')==='ctrTitle')){
      let s = '';
      deep(sp,'a:t').forEach(t => s += t.textContent);
      return s.trim();
    }
  }
  return '';
}

// ============================================================
//  排版：把幻灯片 + 备注排进 A4 页
// ============================================================
function buildLayout(){
  if(!deck) return;
  preview.className = 'preview notes-' + opts.noteSize;
  preview.innerHTML = '';

  const usePdf = opts.picSource === 'pdf' && deck.pdfW > 0;
  const aspect = usePdf ? (deck.pdfW / deck.pdfH) : (deck.renderW / deck.renderH);
  const usableW = (186 - SHEET_PAD_MM*2) * MM;
  const usableH = (273 - SHEET_PAD_MM*2) * MM;
  const N = opts.perPage;

  // 每行高度直接由「每页页数」决定：一页放 N 行，每行 ≈ 可用高度 / N。
  // 这是让「每页页数」真正生效的关键 —— 行高决定每页放几张。
  const ROW_PAD_V = 24;                              // .row 上下内边距合计
  const rowTotalH = Math.max(132, usableH / N - 3);  // 整行盒子高度（含内边距）
  const slideMaxH = rowTotalH - ROW_PAD_V;           // 留给幻灯片缩略图的最大高度

  // 在「行高」与「宽度上限」两个约束内按比例求缩略图尺寸
  let thumbW, thumbH, stackColW = 0;
  if(opts.layout === 'side'){
    const widthFrac = { 2:0.66, 3:0.60, 4:0.56, 5:0.53 }[N] || 0.58;
    const maxW = usableW * widthFrac;
    thumbH = slideMaxH;
    thumbW = thumbH * aspect;
    if(thumbW > maxW){ thumbW = maxW; thumbH = thumbW / aspect; }
  } else {
    // 上下结构：内容收进一个居中的栏，留出左右页边距（类似 Word 文档）
    stackColW = usableW * 0.82;
    // 幻灯片在上、备注在下共享行高：每页放得越多，幻灯片相应越小，给备注留出空间
    const hFrac = { 2:0.66, 3:0.44, 4:0.36, 5:0.30 }[N] || 0.50;
    thumbH = slideMaxH * hFrac;
    thumbW = thumbH * aspect;
    if(thumbW > stackColW){ thumbW = stackColW; thumbH = thumbW / aspect; }
  }
  const scale = thumbW / deck.renderW;

  // 生成每一行
  const rows = [];
  deck.slides.forEach(slide => {
    const hasNotes = slide.notes.some(t => t.trim());
    if(opts.hideEmpty && !hasNotes) return;
    rows.push(buildRow(slide, thumbW, thumbH, scale, rowTotalH, stackColW, usePdf, hasNotes));
  });

  if(!rows.length){
    const tip = document.createElement('p');
    tip.style.cssText = 'color:#8a93a6;font-size:14px;padding:40px;';
    tip.textContent = '没有可显示的幻灯片（可能都被「隐藏无备注的页」过滤掉了）。';
    preview.appendChild(tip);
    return;
  }
  paginate(rows, usableH, rowTotalH);
  updateMeta(rows.length);
}

// 某一行备注过长时，自动缩小该行备注字号，使其回到目标行高
function fitRowNotes(row, rowTotalH){
  const txt = row.querySelector('.notes-text');
  if(!txt) return;
  row.style.minHeight = '0px';                       // 量「自然高度」
  let fs = parseFloat(getComputedStyle(txt).fontSize) || 13;
  let guard = 0;
  while(row.getBoundingClientRect().height > rowTotalH + 1 && fs > 8 && guard < 44){
    fs = Math.max(8, fs - 0.5);
    txt.style.fontSize = fs + 'px';
    guard++;
  }
  row.style.minHeight = rowTotalH + 'px';
}

function buildRow(slide, thumbW, thumbH, scale, rowTotalH, stackColW, usePdf, hasNotes){
  const row = document.createElement('div');
  row.className = 'row' + (opts.layout === 'stack' ? ' layout-stack' : '');
  row.style.minHeight = rowTotalH + 'px';

  // 幻灯片缩略图
  const cellSlide = document.createElement('div');
  cellSlide.className = 'cell-slide';
  const thumb = document.createElement('div');
  thumb.className = 'thumb';
  thumb.style.width = thumbW + 'px';
  thumb.style.height = thumbH + 'px';
  if(usePdf && slide.pdfImg){
    const img = document.createElement('img');
    img.src = slide.pdfImg;
    img.style.cssText = 'width:100%;height:100%;display:block;';
    thumb.appendChild(img);
  } else {
    const inner = document.createElement('div');
    inner.className = 'thumb-inner';
    inner.style.transform = 'scale(' + scale + ')';
    inner.appendChild(slide.domEl.cloneNode(true));
    thumb.appendChild(inner);
  }
  cellSlide.appendChild(thumb);

  // 备注卡片
  const cellNotes = document.createElement('div');
  cellNotes.className = 'cell-notes';
  if(opts.layout === 'stack' && stackColW){
    cellNotes.style.width = stackColW + 'px';   // 居中内容栏，留出左右页边距
  }
  const card = document.createElement('div');
  card.className = 'notes-card' + (hasNotes ? '' : ' is-empty');
  if(opts.showNum || slide.title){
    const head = document.createElement('div');
    head.className = 'note-head';
    if(opts.showNum){
      const badge = document.createElement('span');
      badge.className = 'note-pageno';
      badge.textContent = '第 ' + slide.index + ' 页';
      head.appendChild(badge);
    }
    if(slide.title){
      const tt = document.createElement('span');
      tt.className = 'note-title';
      tt.textContent = slide.title;
      head.appendChild(tt);
    }
    card.appendChild(head);
  }
  const notesText = document.createElement('div');
  notesText.className = 'notes-text';
  if(hasNotes){
    slide.notes.forEach(para => {
      const d = document.createElement('div');
      d.className = 'note-para';
      d.textContent = (para && para.trim()) ? para : '​';
      notesText.appendChild(d);
    });
  } else {
    const e = document.createElement('div');
    e.className = 'note-empty';
    e.textContent = '（此页没有备注）';
    notesText.appendChild(e);
  }
  card.appendChild(notesText);
  cellNotes.appendChild(card);

  row.appendChild(cellSlide);
  row.appendChild(cellNotes);
  return row;
}

function paginate(rows, usableH, rowTotalH){
  // 先在隐藏容器里量一遍，把过长的备注自动缩小，让每行尽量回到目标行高
  const measure = document.createElement('div');
  measure.className = 'sheet';
  measure.style.cssText = 'position:absolute;left:-10000px;top:0;height:auto;visibility:hidden;';
  preview.appendChild(measure);
  rows.forEach(r => measure.appendChild(r));
  rows.forEach(r => fitRowNotes(r, rowTotalH));
  preview.removeChild(measure);

  let sheet = null, used = 0;
  const sheets = [];
  const newSheet = () => {
    sheet = document.createElement('div');
    sheet.className = 'sheet';
    preview.appendChild(sheet);
    sheets.push(sheet);
    used = 0;
  };
  newSheet();
  rows.forEach(row => {
    sheet.appendChild(row);
    let h = row.getBoundingClientRect().height;
    if(used > 0 && used + h > usableH){
      newSheet();
      sheet.appendChild(row);
      h = row.getBoundingClientRect().height;
    }
    used += h;
  });
  sheets.forEach((s, i) => {
    const label = document.createElement('div');
    label.className = 'page-label';
    label.textContent = '第 ' + (i+1) + ' / ' + sheets.length + ' 页';
    s.appendChild(label);
  });
}

function updateMeta(rowCount){
  const total = deck.slides.length;
  const sheets = preview.querySelectorAll('.sheet').length;
  meta.textContent = `共 ${total} 张幻灯片`
    + (rowCount !== total ? `（显示 ${rowCount} 张）` : '')
    + `，排成 ${sheets} 页 A4。`;
}

// ============================================================
//  交互
// ============================================================
function showStatus(msg, kind){
  statusEl.hidden = false;
  statusEl.className = 'status ' + (kind || 'error');
  statusEl.textContent = msg;
}
function clearStatus(){ statusEl.hidden = true; }
function setLoading(text){
  const sp = loading.querySelector('span');
  if(sp) sp.textContent = text || '正在处理…';
  loading.hidden = false;
}

// 用 pdf.js 把 PDF 每页渲染成图片
async function renderPdfToImages(buffer){
  if(typeof pdfjsLib === 'undefined')
    throw new Error('PDF 渲染库未加载（请通过本地服务器打开，或确认 lib/pdf.min.js 存在）。');
  const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
  const pages = [];
  for(let i = 1; i <= pdf.numPages; i++){
    const page = await pdf.getPage(i);
    const base = page.getViewport({ scale: 1 });
    const vp = page.getViewport({ scale: 1600 / base.width });
    const canvas = document.createElement('canvas');
    canvas.width = Math.round(vp.width);
    canvas.height = Math.round(vp.height);
    await page.render({ canvasContext: canvas.getContext('2d'), viewport: vp }).promise;
    pages.push({ src: canvas.toDataURL('image/jpeg', 0.86), w: base.width, h: base.height });
  }
  return pages;
}

// 把 PDF 画面套用到当前 deck
async function applyPdf(file){
  if(!deck){ showStatus('请先加载 PPT 文件。'); return; }
  setLoading('正在渲染 PDF 画面…');
  try {
    const pages = await renderPdfToImages(await file.arrayBuffer());
    if(!pages.length) throw new Error('这个 PDF 里没有页面。');
    deck.slides.forEach((s, i) => { s.pdfImg = pages[i] ? pages[i].src : null; });
    deck.pdfW = pages[0].w;
    deck.pdfH = pages[0].h;
    opts.picSource = 'pdf';
    enablePdfUI();
    if(pages.length !== deck.slides.length){
      showStatus(`提示：PDF 有 ${pages.length} 页，PPT 有 ${deck.slides.length} 张，数量不一致；`
        + '已按先后顺序配对。若 PPT 里有隐藏页，导出 PDF 时请不要跳过隐藏页。', 'warn');
    } else {
      clearStatus();
    }
    buildLayout();
  } catch(err){
    console.error(err);
    showStatus('PDF 加载失败：' + (err.message || '请确认文件无误。'));
  } finally {
    loading.hidden = true;
  }
}
function enablePdfUI(){
  const pdfBtnEl = document.querySelector('.seg[data-control="picSource"] [data-value="pdf"]');
  if(pdfBtnEl) pdfBtnEl.disabled = false;
  document.querySelectorAll('.seg[data-control="picSource"] button')
    .forEach(b => b.classList.toggle('active', b.dataset.value === 'pdf'));
  $('pdfBtn').textContent = '✓ 已加载 PDF（点此更换）';
  $('picTip').textContent = '幻灯片画面来自 PDF，与原文件完全一致。';
}
function resetPdfUI(){
  opts.picSource = 'render';
  const pdfBtnEl = document.querySelector('.seg[data-control="picSource"] [data-value="pdf"]');
  if(pdfBtnEl) pdfBtnEl.disabled = true;
  document.querySelectorAll('.seg[data-control="picSource"] button')
    .forEach(b => b.classList.toggle('active', b.dataset.value === 'render'));
  $('pdfBtn').textContent = '＋ 加载导出的 PDF（画面更精确）';
  $('picTip').textContent = '「自动渲染」可能与原 PPT 有偏差；加载同一份 PPT 导出的 PDF，可让画面精确还原。';
  $('pdfInput').value = '';
}

// 入口：可接收 PPT、PDF，或两者
async function handleFiles(fileList){
  const files = Array.from(fileList || []);
  const pptx = files.find(f => /\.pptx$/i.test(f.name));
  const pdf  = files.find(f => /\.pdf$/i.test(f.name));
  if(!pptx){
    if(pdf) showStatus('还需要 .pptx 文件来读取备注。请把 PPT 和 PDF 一起选择或拖入。');
    else showStatus('请选择 .pptx 文件。旧版 .ppt 请先在 PowerPoint 里「另存为」成 .pptx。');
    return;
  }
  clearStatus();
  setLoading('正在解析 PPT…');
  try {
    deck = await loadPptx(await pptx.arrayBuffer(), pptx.name);
    dropzone.hidden = true;
    workspace.hidden = false;
    resetPdfUI();
    buildLayout();
    if(deck.slides.every(s => !s.notes.some(t => t.trim()))){
      showStatus('提示：这份 PPT 里没有检测到任何备注/演讲者注释，仍可打印幻灯片排版。', 'warn');
    }
  } catch(err){
    console.error(err);
    showStatus(err.message || '解析失败，请确认文件无误。');
    loading.hidden = true;
    return;
  }
  loading.hidden = true;
  if(pdf) await applyPdf(pdf);
}

function resetAll(){
  deck = null;
  workspace.hidden = true;
  dropzone.hidden = false;
  preview.innerHTML = '';
  fileInput.value = '';
  resetPdfUI();
  clearStatus();
}

// 按需加载内置示例数据（用 <script> 注入，双击打开 file:// 也能用，不依赖 fetch）
function loadScriptOnce(src){
  return new Promise((resolve, reject) => {
    if(window.__SAMPLE){ resolve(); return; }
    const s = document.createElement('script');
    s.src = src;
    s.onload = () => resolve();
    s.onerror = () => reject(new Error('无法加载示例数据文件 ' + src));
    document.head.appendChild(s);
  });
}
function b64ToBuffer(b64){
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for(let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes.buffer;
}
async function loadSample(){
  clearStatus();
  setLoading('正在加载示例文件…');
  try {
    await loadScriptOnce('sample-data.js');
    if(!window.__SAMPLE || !window.__SAMPLE.pptx) throw new Error('示例数据缺失。');
    deck = await loadPptx(b64ToBuffer(window.__SAMPLE.pptx), 'sample.pptx');
    dropzone.hidden = true;
    workspace.hidden = false;
    resetPdfUI();
    buildLayout();
  } catch(err){
    console.error(err);
    showStatus('示例文件加载失败：' + (err.message || '请改用「选择文件」拖入自己的 PPT。'), 'warn');
    loading.hidden = true;
    return;
  }
  loading.hidden = true;
  // 示例自带导出的 PDF，演示「PDF 精确画面」
  try {
    if(window.__SAMPLE.pdf)
      await applyPdf(new File([b64ToBuffer(window.__SAMPLE.pdf)], 'sample.pdf'));
  } catch(e){ console.warn(e); }
}

// 控件
function bindControls(){
  document.querySelectorAll('.seg').forEach(seg => {
    seg.addEventListener('click', e => {
      const btn = e.target.closest('button');
      if(!btn) return;
      seg.querySelectorAll('button').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      const ctrl = seg.dataset.control, val = btn.dataset.value;
      opts[ctrl] = (ctrl === 'perPage') ? parseInt(val,10) : val;
      buildLayout();
    });
  });
  $('showNum').addEventListener('change', e => { opts.showNum = e.target.checked; buildLayout(); });
  $('hideEmpty').addEventListener('change', e => { opts.hideEmpty = e.target.checked; buildLayout(); });
  $('printBtn').addEventListener('click', () => window.print());
  $('resetBtn').addEventListener('click', resetAll);
}

// 上传交互
function bindUpload(){
  $('pickBtn').addEventListener('click', () => fileInput.click());
  $('sampleBtn').addEventListener('click', loadSample);
  fileInput.addEventListener('change', e => handleFiles(e.target.files));
  ['dragenter','dragover'].forEach(ev =>
    dropzone.addEventListener(ev, e => { e.preventDefault(); dropzone.classList.add('drag'); }));
  ['dragleave','drop'].forEach(ev =>
    dropzone.addEventListener(ev, e => { e.preventDefault(); dropzone.classList.remove('drag'); }));
  dropzone.addEventListener('drop', e => {
    if(e.dataTransfer.files && e.dataTransfer.files.length) handleFiles(e.dataTransfer.files);
  });
  // 工作区内追加 / 更换 PDF
  $('pdfBtn').addEventListener('click', () => $('pdfInput').click());
  $('pdfInput').addEventListener('change', e => {
    if(e.target.files && e.target.files[0]) applyPdf(e.target.files[0]);
  });
}

bindControls();
bindUpload();

})();
