// @ts-nocheck
// Portado do gerador offline "Kairos - Indicadores Data Quality".
// A lógica de leitura das planilhas e de preenchimento do template PPTX é
// mantida idêntica ao gerador original: o arquivo final usa exatamente o
// mesmo template (template.pptx), apenas com os valores substituídos.
/* eslint-disable @typescript-eslint/no-explicit-any, prefer-const, no-empty */
import JSZip from "jszip";
import * as XLSX from "xlsx";
import templateAsset from "@/assets/template.pptx.asset.json";

export type LogKind = "ok" | "warn" | "err" | undefined;
export type GeneratorFiles = {
  fv: File; fe: File; ac: File; as: File; sm: File;
};

let log: (msg: string, cls?: LogKind) => void = () => {};

const templatePublicUrl = `${import.meta.env.BASE_URL}template.pptx`;
const TEMPLATE_URLS = [templatePublicUrl, templateAsset.url];
async function loadTemplateBuffer(): Promise<ArrayBuffer> {
  let lastStatus = 0;
  const attemptedUrls: string[] = [];
  for (const url of TEMPLATE_URLS) {
    try {
      const resolvedUrl = new URL(url, window.location.href).href;
      attemptedUrls.push(resolvedUrl);
      const response = await fetch(resolvedUrl, { cache: "no-store" });
      if (response.ok) return await response.arrayBuffer();
      lastStatus = response.status;
    } catch {
      lastStatus = 0;
    }
  }
  throw new Error(
    `Não foi possível carregar o template. Status: ${lastStatus}. URLs testadas: ${attemptedUrls.join(", ")}.`,
  );
}

const MESES_PT = ['janeiro','fevereiro','março','abril','maio','junho','julho','agosto','setembro','outubro','novembro','dezembro'];
const cap = s => s.charAt(0).toUpperCase()+s.slice(1);
const fmtPctBR = v => (v*100).toFixed(1).replace('.',',')+'%';    // 0.611 -> "61,1%"
const fmtNumBR = v => v.toFixed(1).replace('.',',');              // 61.1 -> "61,1"

// ===== Detecção de linhas riscadas (strikethrough) =====
// O SheetJS não expõe a fonte "tachado"; lemos o XML do .xlsx direto com JSZip.
async function detectStruckRows(buf){
  const out = {};
  const outCells = {};
  const strikeRe = /<strike\b(?![^>]*val="(?:0|false)")/;
  try {
    const zip = await JSZip.loadAsync(buf);
    const stylesFile = zip.file('xl/styles.xml');
    const styles = stylesFile ? await stylesFile.async('string') : '';

    // 1) fontes tachadas
    const struckFonts = new Set();
    const fontsBlock = (styles.match(/<fonts[\s\S]*?<\/fonts>/) || [''])[0];
    (fontsBlock.match(/<font\b[\s\S]*?(?:\/>|<\/font>)/g) || []).forEach((f, i) => {
      if (strikeRe.test(f)) struckFonts.add(i);
    });

    // 2) cellXfs -> estilos tachados
    const struckXfs = new Set();
    const xfsBlock = (styles.match(/<cellXfs[\s\S]*?<\/cellXfs>/) || [''])[0];
    (xfsBlock.match(/<xf\b[\s\S]*?(?:\/>|<\/xf>)/g) || []).forEach((x, i) => {
      const m = x.match(/fontId="(\d+)"/);
      if (m && struckFonts.has(Number(m[1]))) struckXfs.add(i);
    });

    // 3) shared strings com runs tachados (texto riscado dentro da célula)
    const struckSst = new Set();
    const sstFile = zip.file('xl/sharedStrings.xml');
    if (sstFile) {
      const sst = await sstFile.async('string');
      const items = sst.match(/<si\b[\s\S]*?(?:\/>|<\/si>)/g) || [];
      items.forEach((si, i) => { if (strikeRe.test(si)) struckSst.add(i); });
    }

    if (!struckXfs.size && !struckSst.size) return { rows: out, cells: outCells };

    // nome da aba -> arquivo xml
    const wbXml = await zip.file('xl/workbook.xml').async('string');
    const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');
    const rels = {};
    (relsXml.match(/<Relationship\b[^>]*?(?:\/>|<\/Relationship>)/g) || []).forEach(r => {
      const id = (r.match(/Id="([^"]+)"/) || [])[1];
      const tgt = (r.match(/Target="([^"]+)"/) || [])[1];
      if (id && tgt) rels[id] = tgt.replace(/^\/?xl\//, '').replace(/^\.\//, '');
    });
    const sheetTags = wbXml.match(/<sheet\b[^>]*?(?:\/>|<\/sheet>)/g) || [];
    for (const st of sheetTags) {
      const name = (st.match(/name="([^"]*)"/) || [])[1];
      const rid = (st.match(/r:id="([^"]+)"/) || [])[1];
      const path = rid && rels[rid] ? 'xl/' + rels[rid] : null;
      if (!name || !path || !zip.file(path)) continue;
      const xml = await zip.file(path).async('string');

      // colunas com estilo tachado (célula sem s= herda o estilo da coluna)
      const struckCols = new Set();
      (xml.match(/<col\b[^>]*\/>/g) || []).forEach(c => {
        const sty = (c.match(/\bstyle="(\d+)"/) || [])[1];
        if (sty == null || !struckXfs.has(Number(sty))) return;
        const min = Number((c.match(/\bmin="(\d+)"/) || [])[1] || 0);
        const max = Number((c.match(/\bmax="(\d+)"/) || [])[1] || min);
        for (let k = min; k <= max && k - min < 16384; k++) struckCols.add(k);
      });

      const set = new Set();
      const cellSet = new Set();
      const rowRe = /<row\b([^>]*)(\/>|>([\s\S]*?)<\/row>)/g;
      let m;
      while ((m = rowRe.exec(xml)) !== null) {
        const attrs = m[1] || '';
        const body = m[3] || '';
        const rNum = Number((attrs.match(/\br="(\d+)"/) || [])[1]);
        if (!rNum) continue;
        let rowStruck = false;
        // estilo aplicado a linha inteira
        const rs = attrs.match(/\bs="(\d+)"/);
        if (rs && /customFormat="(?:1|true)"/.test(attrs) && struckXfs.has(Number(rs[1]))) rowStruck = true;
        let content = 0, struckContent = 0;
        if (body) {
          const cells = body.match(/<c\b[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g) || [];
          for (const c of cells) {
            const ref = (c.match(/\br="([A-Z]+)\d+"/) || [])[1];
            const hasVal = /<v>|<is>|<t>/.test(c);
            const cs = (c.match(/\bs="(\d+)"/) || [])[1];
            let cStruck = false;
            if (cs != null && struckXfs.has(Number(cs))) cStruck = true;
            if (!cStruck && cs == null && ref && struckCols.has(colIdx(ref))) cStruck = true;
            if (!cStruck && /\bt="s"/.test(c)) {
              const v = (c.match(/<v>(\d+)<\/v>/) || [])[1];
              if (v != null && struckSst.has(Number(v))) cStruck = true;
            }
            if (!cStruck && /\bt="inlineStr"/.test(c) && strikeRe.test(c)) cStruck = true;
            if (hasVal) content++;
            if (cStruck) {
              if (hasVal) struckContent++;
              if (ref) cellSet.add((rNum - 1) + ',' + (colIdx(ref) - 1)); // celula riscada (0-based)
            }
          }
        }
        // linha riscada apenas quando o estilo e da linha inteira ou todas as celulas com conteudo estao riscadas
        if (rowStruck || (content > 0 && struckContent === content)) set.add(rNum - 1);
      }
      if (set.size) out[name] = set;
      if (cellSet.size) outCells[name] = cellSet;
    }
  } catch (e) { /* silencioso: sem detecção, nada é excluído */ }
  return { rows: out, cells: outCells };
}
// A1 -> índice de coluna 1-based
function colIdx(letters){
  let n = 0;
  for (let i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
  return n;
}

// Conteúdo literalmente tachado em Markdown: espaços externos são permitidos,
// mas deve existir conteúdo entre os pares de ~~.
function isMarkdownStruck(value){
  if (value == null) return false;
  const text = String(value).trim();
  return /^~~[\s\S]+~~$/.test(text) && text.slice(2, -2).length > 0;
}

// Detecta o tachado Markdown antes de qualquer parser de negócio.
// Uma linha é integralmente riscada quando todas as suas células preenchidas
// estão em ~~...~~; em linhas válidas, somente as células marcadas são anuladas.
function mergeMarkdownStruck(wb, rowsBySheet, cellsBySheet){
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    if (!ws || !ws['!ref']) continue;
    const range = XLSX.utils.decode_range(ws['!ref']);
    const grid = XLSX.utils.sheet_to_json(ws, { header:1, defval:null, raw:false });
    const rowSet = rowsBySheet[name] || new Set();
    const cellSet = cellsBySheet[name] || new Set();
    for (let i = 0; i < grid.length; i++) {
      const row = grid[i] || [];
      let filled = 0, markdownStruck = 0;
      for (let j = 0; j < row.length; j++) {
        const value = row[j];
        if (value == null || String(value).trim() === '') continue;
        filled++;
        if (isMarkdownStruck(value)) {
          markdownStruck++;
          cellSet.add((range.s.r + i) + ',' + (range.s.c + j));
        }
      }
      if (filled > 0 && markdownStruck === filled) rowSet.add(range.s.r + i);
    }
    if (rowSet.size) rowsBySheet[name] = rowSet;
    if (cellSet.size) cellsBySheet[name] = cellSet;
  }
}

// Conjunto de linhas riscadas da aba (índices absolutos 0-based)
function struckSet(wb, sheetName){
  return (wb && wb.__struck && wb.__struck[sheetName]) || null;
}
// base = primeira linha do range da aba (0-based)
function sheetBase(ws){
  try { return XLSX.utils.decode_range(ws['!ref']).s.r; } catch(e){ return 0; }
}
function totalStruck(wb){
  let n = 0;
  if (wb && wb.__struck) for (const k in wb.__struck) n += wb.__struck[k].size;
  return n;
}

async function readXlsx(file){
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type:'array', cellDates:true });
  const det = await detectStruckRows(buf);
  // Une tachado de formatação do Excel e tachado Markdown literal.
  // Esta etapa acontece antes da remoção lógica de células e de todos os cálculos.
  mergeMarkdownStruck(wb, det.rows, det.cells);
  wb.__struck = det.rows;
  wb.__struckCells = det.cells;
  // Ordem correta: linhas riscadas ficam fora e celulas riscadas sao removidas
  // logicamente ANTES de qualquer identificacao de ofensores.
  for (const nm in det.cells) {
    const ws = wb.Sheets[nm];
    if (!ws) continue;
    det.cells[nm].forEach(k => {
      const p = k.split(',');
      const addr = XLSX.utils.encode_cell({ r: Number(p[0]), c: Number(p[1]) });
      if (ws[addr]) delete ws[addr];
    });
  }
  return wb;
}



// Score mensal
function parseScoreMensal(wb){
  const name = wb.SheetNames[0];
  const ws = wb.Sheets[name];
  const struck = struckSet(wb, name), base = sheetBase(ws);
  const grid = XLSX.utils.sheet_to_json(ws, { header:1, defval:null, raw:true });
  const rows=[];
  for (let i=0;i<grid.length;i++) {
    if (struck && struck.has(base + i)) continue; // linha riscada: ignorada
    const r = grid[i];
    if (r && r[0] instanceof Date && typeof r[1]==='number') rows.push({ date:r[0], score:r[1] });
  }
  rows.sort((a,b)=>a.date-b.date);
  return rows;
}

// Validação mensal — col C = nome, col D = CNPJ, col E = status
function parseValidacao(wb){
  const name = wb.SheetNames[0];
  const ws = wb.Sheets[name];
  const struck = struckSet(wb, name), base = sheetBase(ws);
  const isStruck = i => !!(struck && struck.has(base + i));
  const grid = XLSX.utils.sheet_to_json(ws, { header:1, defval:null, raw:false });
  let latest=null;
  for (let i=1;i<grid.length;i++){
    if (isStruck(i)) continue; // riscada: não define o mês de referência
    const m = grid[i] && grid[i][0];
    if (typeof m==='string' && /^\d{2}-\d{2}$/.test(m.trim())) {
      const v=m.trim(); if (!latest || v>latest) latest=v;
    }
  }
  let dentro=0, fora=0, pend=0, div=0, total=0;
  const tabela=[];
  let riscadas=0;
  for (let i=1;i<grid.length;i++){
    if (isStruck(i)) { riscadas++; continue; } // riscada: fora de todos os status
    const r = grid[i]; if(!r) continue;
    const mes = typeof r[0]==='string' ? r[0].trim() : r[0];
    if (mes !== latest) continue;
    const nome = r[2]!=null ? String(r[2]).trim() : '';
    const cnpj = r[3]!=null ? String(r[3]).trim() : '';
    const st = r[4]!=null ? String(r[4]).trim() : '';
    if (!cnpj && !nome) continue;
    total++;
    const low = st.toLowerCase();
    if (low.includes('diverg')) { div++; tabela.push({cnpj, nome, status:'Divergência'}); }
    else if (low.includes('pend')) { pend++; tabela.push({cnpj, nome, status:'Pendente'}); }
    else if (low.includes('fora')) fora++;
    else if (low.includes('dentro') || low.includes('valid')) dentro++;
  }
  tabela.sort((a,b)=> a.status===b.status ? a.nome.localeCompare(b.nome) : a.status.localeCompare(b.status));
  const d = total || 1;
  return { latest, total, riscadas, dentro, fora, pend, div, tabela,
    pctDentro: dentro/d, pctFora: fora/d, pctPend: pend/d, pctDiv: div/d,
    pctValidado: (dentro+fora)/d };
}

// Frequência (Vendas ou Estoque) — conta dias "I" (indisponível / vermelho)
function parseFrequenciaDias(wb){
  const name = wb.SheetNames[0];
  const ws = wb.Sheets[name];
  const struck = struckSet(wb, name), base = sheetBase(ws);
  const grid = XLSX.utils.sheet_to_json(ws, { header:1, defval:null, raw:false });
  let mesLabel=null;
  const filt = grid[0] && grid[0][0];
  if (typeof filt==='string'){
    const m = filt.match(/Mês\s*=\s*([A-Za-zç]+\/\d{4})/);
    if (m) mesLabel=m[1];
  }
  // localiza a coluna de % Frequência existente na propria planilha (cabecalho)
  let colFreq = -1;
  for (let i=0;i<Math.min(6, grid.length) && colFreq<0;i++){
    const r = grid[i]; if (!r) continue;
    for (let j=0;j<Math.min(r.length,6);j++){
      const v = r[j]; if (v==null) continue;
      const t = String(v).toLowerCase();
      if (t.includes('frequ') || t.includes('%')) { colFreq=j; break; }
    }
  }
  const parsePct = v => {
    if (v==null) return null;
    if (typeof v==='number') return v>1 ? v/100 : v;
    let t = String(v).trim();
    if (!t || isMarkdownStruck(t)) return null;
    const hasPct = t.includes('%');
    t = t.replace('%','').replace(/\s/g,'').replace(/\./g,'').replace(',','.');
    const n = parseFloat(t);
    if (isNaN(n)) return null;
    return (hasPct || n>1) ? n/100 : n;
  };

  const byCnpj = new Map();
  const isStruck = i => !!(struck && struck.has(base + i));
  for (let i=5;i<grid.length;i++){
    if (isStruck(i)) continue; // riscada: fora do cálculo
    const r=grid[i]; if(!r) continue;
    // mesma regra da validação mensal: linha integralmente tachada é ignorada
    let filled=0, struckCells=0;
    for (let j=0;j<r.length;j++){
      const v=r[j];
      if (v==null || String(v).trim()==='') continue;
      filled++;
      if (isMarkdownStruck(v)) struckCells++;
    }
    if (filled>0 && struckCells===filled) continue;
    const nome = !isMarkdownStruck(r[0]) && r[0]!=null ? String(r[0]).trim() : '';
    const cnpj = !isMarkdownStruck(r[1]) && r[1]!=null ? String(r[1]).trim() : '';
    if (!cnpj) continue;
    let ausentes=0, considerados=0;
    for (let j=5;j<r.length;j++){
      const v = r[j]; if (v==null) continue;
      if (isMarkdownStruck(v)) continue;      // célula riscada: ignorada
      const s = String(v).trim().toUpperCase();
      if (!s) continue;
       
      considerados++;
      if (s==='I') ausentes++;
    }
    if (considerados===0) continue;
    const pctPlan = colFreq>=0 ? parsePct(r[colFreq]) : null;
    const prev = byCnpj.get(cnpj);
    if (!prev || ausentes > prev.ausentes) byCnpj.set(cnpj, { cnpj, nome, ausentes, considerados, pctPlan });
  }
  const list = [...byCnpj.values()];
  let atualizados=0, atencao=0, ofensores=0, somaAus=0, somaCon=0;
  for (const d of list){
    if (d.ausentes>=5) ofensores++;
    else if (d.ausentes>=1) atencao++;
    else atualizados++;
    somaAus += d.ausentes; somaCon += d.considerados;
  }
  // % Frequência: usa a da planilha quando existir; senão, calcula
  const pctFreqRow = d => d.pctPlan != null ? Number(d.pctPlan)
    : (d.considerados ? (Number(d.considerados)-Number(d.ausentes))/Number(d.considerados) : 0);
  const pctAus = d => 1 - pctFreqRow(d);
  // ordenação única: % de dias sem dados (maior → menor), depois qtd de dias e nome
  const ofensoresList = list.filter(d=>d.ausentes>=5)
    .sort((a,b)=> (pctAus(b)-pctAus(a)) || (Number(b.ausentes)-Number(a.ausentes)) || a.nome.localeCompare(b.nome));


  const pctFreq = somaCon ? (somaCon-somaAus)/somaCon : 0;
  return { mesLabel, list, atualizados, atencao, ofensores, total:list.length, ofensoresList, pctFreq };
}

// Associações
function parseAssoc(wb){
  const resumoName = wb.Sheets['Resumo'] ? 'Resumo' : wb.SheetNames[0];
  const resumoSheet = wb.Sheets[resumoName];
  const resumoStruck = struckSet(wb, resumoName), resumoBase = sheetBase(resumoSheet);
  const resumoGrid = XLSX.utils.sheet_to_json(resumoSheet, { header:1, defval:null });
  let dist=0, ind=0;
  for (let ri=0; ri<resumoGrid.length; ri++) {
    if (resumoStruck && resumoStruck.has(resumoBase + ri)) continue;
    const r = resumoGrid[ri];
    if (!r) continue;
    const nums = r.filter(x=>typeof x==='number');
    if (nums.length>=3) { ind = Math.round(nums[1]); dist = Math.round(nums[2]); break; }
  }
  const itemsName = wb.Sheets['Associação de Itens'] ? 'Associação de Itens' : wb.SheetNames[1];
  const itemsSheet = wb.Sheets[itemsName];
  const itemsStruck = itemsSheet ? struckSet(wb, itemsName) : null;
  const itemsBase = itemsSheet ? sheetBase(itemsSheet) : 0;
  let sim=0, nao=0, rejeit=0;
  if (itemsSheet) {
    const grid = XLSX.utils.sheet_to_json(itemsSheet, { header:1, defval:null });
    let hdr=-1, cAssoc=-1;
    for (let i=0;i<grid.length;i++){
      const r=grid[i]; if(!r) continue;
      const idx = r.findIndex(x=>typeof x==='string' && x.trim()==='Item foi Associado?');
      if (idx>=0){ hdr=i; cAssoc=idx; break; }
    }
    if (hdr>=0){
      for (let i=hdr+1;i<grid.length;i++){
        if (itemsStruck && itemsStruck.has(itemsBase + i)) continue; // riscada: fora do cálculo
        const r=grid[i]; if(!r) continue;
        const v = r[cAssoc];
        if (v==null || v==='') { if (r.some(x=>x!=null && x!=='')) rejeit++; }
        else { const s=String(v).trim().toUpperCase(); if(s==='SIM') sim++; else if(s==='NÃO'||s==='NAO') nao++; }
      }
    }
  }
  return { distribuidores: dist, industrias: ind, associados: sim, naoIdent: nao, rejeitados: rejeit, impactados: sim };
}

function xmlEscape(s){ return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }

function replaceRun(xml: string, oldText, newText){
  const target = '<a:t>'+oldText+'</a:t>';
  const rep = '<a:t>'+xmlEscape(newText)+'</a:t>';
  const i = xml.indexOf(target);
  if (i<0) { log('  ! não encontrado: '+oldText, 'warn'); return xml; }
  return xml.slice(0,i)+rep+xml.slice(i+target.length);
}

function replaceLastRun(xml, oldText, newText){
  const target = '<a:t>'+oldText+'</a:t>';
  const i = xml.lastIndexOf(target);
  if (i<0) { log('  ! não encontrado (último): '+oldText, 'warn'); return xml; }
  return xml.slice(0,i)+'<a:t>'+xmlEscape(newText)+'</a:t>'+xml.slice(i+target.length);
}

// Substitui os N primeiros runs cujo texto é igual a oldText, na ordem
function replaceRunsSeq(xml, oldText, values){
  let out = xml, cursor = 0, from = 0;
  const target = '<a:t>'+oldText+'</a:t>';
  while (cursor < values.length) {
    const i = out.indexOf(target, from);
    if (i<0) { log('  ! run não encontrado: '+oldText, 'warn'); break; }
    const rep = '<a:t>'+xmlEscape(values[cursor])+'</a:t>';
    out = out.slice(0,i)+rep+out.slice(i+target.length);
    from = i + rep.length;
    cursor++;
  }
  return out;
}

function countTableDataRows(xml){
  const tblStart = xml.indexOf('<a:tbl>');
  if (tblStart<0) return 0;
  const tblEnd = xml.indexOf('</a:tbl>', tblStart)+8;
  const tbl = xml.slice(tblStart, tblEnd);
  const trs = tbl.match(/<a:tr[ >]/g) || [];
  return Math.max(0, trs.length-1);
}

// Preenche a primeira tabela do slide: rows = array de arrays (sem cabeçalho)
function fillTable(xml, rows){
  const tblStart = xml.indexOf('<a:tbl>');
  if (tblStart<0){ log('  ! tabela não encontrada', 'warn'); return xml; }
  const tblEnd = xml.indexOf('</a:tbl>', tblStart)+8;
  let tbl = xml.slice(tblStart, tblEnd);
  const trRe = /<a:tr[ >][\s\S]*?<\/a:tr>/g;
  let dataIdx = -1;
  tbl = tbl.replace(trRe, (tr) => {
    dataIdx++;
    if (dataIdx===0) return tr;               // cabeçalho
    const data = rows[dataIdx-1];
    // sem dados para esta linha → remove a linha da tabela (não deixa linhas em branco)
    if (!data || !data.some(v => v != null && String(v).trim() !== '')) return '';
    let cellIdx = -1;
    return tr.replace(/<a:tc>[\s\S]*?<\/a:tc>/g, (tc) => {
      cellIdx++;
      const val = data && data[cellIdx] != null ? String(data[cellIdx]) : ' ';
      if (!/<a:t>/.test(tc)) {
        // célula sem run de texto no template → cria um run herdando a formatação do parágrafo
        return tc.replace(/<a:endParaRPr([^>]*)(\/>|>[\s\S]*?<\/a:endParaRPr>)/, (m, attrs, rest) => {
          const inner = rest === '/>' ? '' : rest.slice(1, rest.lastIndexOf('</a:endParaRPr>'));
          const rPr = '<a:rPr'+attrs+(inner ? '>'+inner+'</a:rPr>' : '/>');
          return '<a:r>'+rPr+'<a:t>'+xmlEscape(val)+'</a:t></a:r>'+m;
        });
      }
      let done = false;
      return tc.replace(/<a:t>[\s\S]*?<\/a:t>/g, () => {
        if (done) return '<a:t></a:t>';
        done = true;
        return '<a:t>'+xmlEscape(val)+'</a:t>';
      });
    });

  });

  return xml.slice(0,tblStart)+tbl+xml.slice(tblEnd);
}

// Substitui <c:pt idx="i"><c:v>...</c:v></c:pt> dentro de um trecho
function replacePtInSection(xml, sectionStart, sectionEndTag, idx, newVal){
  const start = xml.indexOf(sectionStart);
  if (start<0) return xml;
  const end = xml.indexOf(sectionEndTag, start);
  if (end<0) return xml;
  const middle = xml.slice(start, end);
  const re = new RegExp('(<c:pt idx="'+idx+'"[^>]*>\\s*<c:v>)[^<]*(</c:v>\\s*</c:pt>)');
  return xml.slice(0,start) + middle.replace(re, '$1'+xmlEscape(newVal)+'$2') + xml.slice(end);
}

// Reescreve TODOS os pontos de um trecho (<c:cat> ou <c:val>) com a lista completa
function setPointsInSection(xml, sectionStart, sectionEndTag, values){
  const start = xml.indexOf(sectionStart);
  if (start<0) return xml;
  const end = xml.indexOf(sectionEndTag, start);
  if (end<0) return xml;
  let middle = xml.slice(start, end);
  const ptRe = /<c:pt\s+idx="\d+"([^>]*)>\s*<c:v>[^<]*<\/c:v>\s*<\/c:pt>/g;
  const first = middle.match(/<c:pt\s+idx="\d+"([^>]*)>/);
  const extraAttrs = first ? first[1] : '';
  const pts = values.map((v,i)=>'<c:pt idx="'+i+'"'+extraAttrs+'><c:v>'+xmlEscape(String(v))+'</c:v></c:pt>').join('');
  // remove pontos existentes
  middle = middle.replace(ptRe, '');
  // atualiza contagem
  middle = middle.replace(/<c:ptCount\s+val="\d+"\s*\/>/, '<c:ptCount val="'+values.length+'"/>');
  // insere os novos pontos após o ptCount (ou no fim do cache)
  const ptCountIdx = middle.indexOf('<c:ptCount');
  if (ptCountIdx >= 0) {
    const insertAt = middle.indexOf('/>', ptCountIdx)+2;
    middle = middle.slice(0, insertAt) + pts + middle.slice(insertAt);
  } else {
    middle += pts;
  }
  return xml.slice(0,start) + middle + xml.slice(end);
}


// Excel date serial (1900-based) from a JS Date
function dateToSerial(d){
  const epoch = new Date(Date.UTC(1899,11,30));
  return Math.round((d - epoch) / 86400000);
}

// ---- Duplicação de slides (para listas longas de ofensores) ----
async function addSlideCopy(zip, srcNum, newXml, afterNum){
  // descobre o próximo número de slide livre
  let n = 1;
  while (zip.file('ppt/slides/slide'+n+'.xml')) n++;
  const newNum = n;
  zip.file('ppt/slides/slide'+newNum+'.xml', newXml);
  let srcRels = await zip.file('ppt/slides/_rels/slide'+srcNum+'.xml.rels').async('string');
  // remove vínculos que não podem ser compartilhados entre slides (anotações e comentários)
  srcRels = srcRels.replace(/<Relationship[^>]*Target="\.\.\/(notesSlides|comments)\/[^"]*"[^>]*\/>/g, '');
  zip.file('ppt/slides/_rels/slide'+newNum+'.xml.rels', srcRels);

  // Content_Types
  let ct = await zip.file('[Content_Types].xml').async('string');
  ct = ct.replace('</Types>',
    '<Override PartName="/ppt/slides/slide'+newNum+'.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>');
  zip.file('[Content_Types].xml', ct);

  // rels da apresentação
  let prels = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
  const ids = [...prels.matchAll(/Id="rId(\d+)"/g)].map(m=>parseInt(m[1],10));
  const newRid = 'rId'+(Math.max(...ids)+1);
  prels = prels.replace('</Relationships>',
    '<Relationship Id="'+newRid+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide'+newNum+'.xml"/></Relationships>');
  // rId do slide após o qual a cópia deve entrar (mantém a ordem das páginas)
  const anchorNum = afterNum || srcNum;
  const srcRelMatch = prels.match(new RegExp('Id="(rId\\d+)"[^>]*Target="slides/slide'+anchorNum+'\\.xml"'));
  zip.file('ppt/_rels/presentation.xml.rels', prels);

  // presentation.xml — insere o sldId imediatamente após o slide âncora
  let pres = await zip.file('ppt/presentation.xml').async('string');
  const sldIds = [...pres.matchAll(/<p:sldId id="\d+" r:id="rId\d+"\/>/g)].map(m=>m[0]);
  const maxId = Math.max(...[...pres.matchAll(/<p:sldId id="(\d+)"/g)].map(m=>parseInt(m[1],10)));
  const newSldId = '<p:sldId id="'+(maxId+1)+'" r:id="'+newRid+'"/>';
  let anchor = null;
  if (srcRelMatch) anchor = sldIds.find(s => s.includes('r:id="'+srcRelMatch[1]+'"'));
  if (anchor) {
    const idx = pres.indexOf(anchor)+anchor.length;
    pres = pres.slice(0,idx)+newSldId+pres.slice(idx);
  } else {
    pres = pres.replace('</p:sldIdLst>', newSldId+'</p:sldIdLst>');
  }
  zip.file('ppt/presentation.xml', pres);
  return newNum;
}

// Preenche a tabela do slide `srcNum` paginando todas as linhas, duplicando o slide quando necessário
// Renomeia o texto da ultima celula do cabecalho da primeira tabela
function setTableHeaderLast(xml, text){
  const tblStart = xml.indexOf('<a:tbl>');
  if (tblStart<0) return xml;
  const tblEnd = xml.indexOf('</a:tbl>', tblStart)+8;
  let tbl = xml.slice(tblStart, tblEnd);
  const trMatch = tbl.match(/<a:tr[ >][\s\S]*?<\/a:tr>/);
  if (!trMatch) return xml;
  let tr = trMatch[0];
  const tcs = tr.match(/<a:tc>[\s\S]*?<\/a:tc>/g);
  if (!tcs || !tcs.length) return xml;
  const lastTc = tcs[tcs.length-1];
  if (!/<a:t>/.test(lastTc)) return xml;
  let done=false;
  const newTc = lastTc.replace(/<a:t>[\s\S]*?<\/a:t>/g, () => {
    if (done) return '<a:t></a:t>';
    done = true;
    return '<a:t>'+xmlEscape(text)+'</a:t>';
  });
  tr = tr.replace(lastTc, newTc);
  tbl = tbl.replace(trMatch[0], tr);
  return xml.slice(0,tblStart)+tbl+xml.slice(tblEnd);
}

async function fillTablePaged(zip, srcNum, rows, label, headerLast){
  let src = await zip.file('ppt/slides/slide'+srcNum+'.xml').async('string');
  if (headerLast) src = setTableHeaderLast(src, headerLast);
  const perPage = countTableDataRows(src) || 29;
  const pages = [];
  for (let i=0;i<rows.length; i+=perPage) pages.push(rows.slice(i, i+perPage));
  if (pages.length===0) pages.push([]);

  zip.file('ppt/slides/slide'+srcNum+'.xml', fillTable(src, pages[0]));
  let lastAnchor = srcNum;
  for (let p=1;p<pages.length;p++){
    const xml = fillTable(src, pages[p]);
    const num = await addSlideCopy(zip, srcNum, xml, lastAnchor);
    lastAnchor = num;
  }
  log('  '+label+': '+rows.length+' linhas em '+pages.length+' slide(s)', 'ok');
}

// ---- Tabelas múltiplas por slide (slide de Validação: 2 colunas) ----
const FRAME_RE = /<p:graphicFrame>[\s\S]*?<\/p:graphicFrame>/g;

function tableFrames(xml){
  const out: { start:number; end:number; text:string }[] = [];
  FRAME_RE.lastIndex = 0;
  let m;
  while ((m = FRAME_RE.exec(xml)) !== null) {
    if (m[0].indexOf('<a:tbl>') >= 0) out.push({ start: m.index, end: m.index + m[0].length, text: m[0] });
  }
  return out;
}

function fillTableNth(xml, idx, rows){
  const frames = tableFrames(xml);
  const f = frames[idx];
  if (!f) return xml;
  return xml.slice(0, f.start) + fillTable(f.text, rows) + xml.slice(f.end);
}

function removeTableNth(xml, idx){
  const frames = tableFrames(xml);
  const f = frames[idx];
  if (!f) return xml;
  return xml.slice(0, f.start) + xml.slice(f.end);
}

// Remove um slide do pacote (arquivo, rels, content types e ordem)
async function removeSlide(zip, num){
  const path = 'ppt/slides/slide'+num+'.xml';
  if (!zip.file(path)) return;
  zip.remove(path);
  zip.remove('ppt/slides/_rels/slide'+num+'.xml.rels');

  let ct = await zip.file('[Content_Types].xml').async('string');
  ct = ct.replace(new RegExp('<Override PartName="/ppt/slides/slide'+num+'\\.xml"[^>]*/>'), '');
  zip.file('[Content_Types].xml', ct);

  let prels = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
  const relMatch = prels.match(new RegExp('<Relationship Id="(rId\\d+)"[^>]*Target="slides/slide'+num+'\\.xml"[^>]*/>'));
  if (relMatch) {
    prels = prels.replace(relMatch[0], '');
    zip.file('ppt/_rels/presentation.xml.rels', prels);
    let pres = await zip.file('ppt/presentation.xml').async('string');
    pres = pres.replace(new RegExp('<p:sldId id="\\d+" r:id="'+relMatch[1]+'"/>'), '');
    zip.file('ppt/presentation.xml', pres);
  }
}

// Pagina uma lista em slides já existentes no template; duplica o último quando
// faltam páginas e remove os slides excedentes quando sobram.
async function fillTableAcrossSlides(zip, slideNums, rows, label, prepare?: (xml:string)=>string){
  const sources: string[] = [];
  for (const n of slideNums) {
    let s = await zip.file('ppt/slides/slide'+n+'.xml').async('string');
    if (prepare) s = prepare(s);
    sources.push(s);
  }
  const perPage = countTableDataRows(sources[0]) || 29;
  const pages: any[][] = [];
  for (let i=0;i<rows.length;i+=perPage) pages.push(rows.slice(i, i+perPage));
  if (!pages.length) pages.push([]);

  const used = Math.min(pages.length, slideNums.length);
  for (let p=0;p<used;p++) zip.file('ppt/slides/slide'+slideNums[p]+'.xml', fillTable(sources[p], pages[p]));
  // páginas extras: cópias do último slide do template (layout de continuação)
  let lastAnchor = slideNums[used-1];
  const lastSrc = sources[sources.length-1];
  for (let p=used;p<pages.length;p++){
    const num = await addSlideCopy(zip, slideNums[slideNums.length-1], fillTable(lastSrc, pages[p]), lastAnchor);
    lastAnchor = num;
  }
  // slides do template que não foram necessários
  for (let p=pages.length;p<slideNums.length;p++) await removeSlide(zip, slideNums[p]);
  log('  '+label+': '+rows.length+' linhas em '+pages.length+' slide(s)', 'ok');
}

// Slide de Validação: duas tabelas lado a lado (12 linhas cada) e duplicação
// do slide inteiro quando a lista ultrapassa a capacidade da página.
async function fillValidacaoSlide(zip, slideNum, rows){
  const src = await zip.file('ppt/slides/slide'+slideNum+'.xml').async('string');
  const frames = tableFrames(src);
  const perTable = countTableDataRows(frames[0] ? frames[0].text : src) || 12;
  const perPage = perTable * Math.max(1, frames.length);
  const pages: any[][] = [];
  for (let i=0;i<rows.length;i+=perPage) pages.push(rows.slice(i, i+perPage));
  if (!pages.length) pages.push([]);

  const build = (pageRows) => {
    let xml = src;
    // preenche da última para a primeira para não invalidar os índices
    for (let t = frames.length-1; t >= 0; t--) {
      const chunk = pageRows.slice(t*perTable, (t+1)*perTable);
      if (t > 0 && chunk.length === 0) xml = removeTableNth(xml, t);
      else xml = fillTableNth(xml, t, chunk);
    }
    return xml;
  };

  zip.file('ppt/slides/slide'+slideNum+'.xml', build(pages[0]));
  let lastAnchor = slideNum;
  for (let p=1;p<pages.length;p++){
    lastAnchor = await addSlideCopy(zip, slideNum, build(pages[p]), lastAnchor);
  }
  log('  Não Validados: '+rows.length+' linhas em '+pages.length+' slide(s) ('+perTable+' por coluna)', 'ok');
}

export async function gerarApresentacao(
  files: GeneratorFiles,
  onLog: (msg: string, cls?: LogKind) => void = () => {},
): Promise<Blob> {
  log = onLog;
  log('Lendo planilhas...');
  const [wbSm, wbAs, wbAc, wbFv, wbFe] = await Promise.all([
    readXlsx(files.sm), readXlsx(files.as),
    readXlsx(files.ac), readXlsx(files.fv), readXlsx(files.fe)
  ]);
  const sm = parseScoreMensal(wbSm);
  const as = parseAssoc(wbAs);
  const val = parseValidacao(wbAc);
  const fv = parseFrequenciaDias(wbFv);
  const fe = parseFrequenciaDias(wbFe);

  const riscTotal = [wbSm,wbAs,wbAc,wbFv,wbFe].reduce((a,w)=>a+totalStruck(w),0);
  log('Linhas riscadas ignoradas nos cálculos: '+riscTotal, riscTotal?'warn':'ok');
  log('Score mensal pontos: '+sm.length, 'ok');
  log('Validação '+val.latest+': dentro '+val.dentro+', fora '+val.fora+', pendentes '+val.pend+', divergentes '+val.div+' (total válido '+val.total+')', 'ok');
  log('Freq Vendas: '+fv.atualizados+' atualizados / '+fv.atencao+' atenção / '+fv.ofensores+' desatualizados (total '+fv.total+')', 'ok');
  log('Freq Estoque: '+fe.atualizados+' atualizados / '+fe.atencao+' atenção / '+fe.ofensores+' desatualizados (total '+fe.total+')', 'ok');
  log('Associações: '+JSON.stringify(as), 'ok');

  let s2mm='07', s2yyyy='2026';
  if (val.latest){ const [yy,mm]=val.latest.split('-'); s2mm=mm; s2yyyy='20'+yy; }
  const smLast = sm[sm.length-1];
  const mesNome = smLast
    ? cap(MESES_PT[smLast.date.getMonth()])+' '+smLast.date.getFullYear()
    : cap(MESES_PT[parseInt(s2mm,10)-1])+' '+s2yyyy;

  log('Carregando template...');
  const zip = await JSZip.loadAsync(await loadTemplateBuffer());

  // ---- Slide 3 — Consolidado ----
  log('Editando Consolidado...');
  let s3 = await zip.file('ppt/slides/slide3.xml').async('string');
  s3 = replaceRun(s3, 'Agosto 2026', mesNome);
  const smRef = sm.find(r => r.date.getMonth()+1 === parseInt(s2mm,10) && r.date.getFullYear() === parseInt(s2yyyy,10)) || smLast;
  const avgFreq = (fv.pctFreq + fe.pctFreq) / 2;
  const avgCons = (as.associados + as.naoIdent) ? as.associados / (as.associados + as.naoIdent) : 1;
  const scoreResumo = smRef ? smRef.score : (avgFreq + val.pctValidado + val.pctDentro + avgCons) / 4;
  s3 = replaceRun(s3, ' 91,2%', ' '+fmtPctBR(scoreResumo));
  s3 = replaceRun(s3, ' 72,8%', ' '+fmtPctBR(avgFreq));
  s3 = replaceRun(s3, ' 96,0%', ' '+fmtPctBR(val.pctValidado));
  s3 = replaceRun(s3, ' 50,0%', ' '+fmtPctBR(val.pctDentro));
  s3 = replaceRun(s3, ' 97,4%', ' '+fmtPctBR(avgCons));
  zip.file('ppt/slides/slide3.xml', s3);

  log('Atualizando gráfico de evolução...');
  let c1 = await zip.file('ppt/charts/chart1.xml').async('string');
  c1 = setPointsInSection(c1, '<c:cat>', '</c:cat>', sm.map(r=>dateToSerial(r.date)));
  c1 = setPointsInSection(c1, '<c:val>', '</c:val>', sm.map(r=>r.score));
  zip.file('ppt/charts/chart1.xml', c1);
  log('Gráfico de evolução: '+sm.length+' meses', 'ok');

  // ---- Slide 4 — Validação Mensal ----
  log('Editando Validação Mensal...');
  let s4 = await zip.file('ppt/slides/slide4.xml').async('string');
  s4 = replaceRun(s4, 'Agosto 2026', mesNome);
  s4 = replaceRun(s4, '96,0', fmtNumBR(val.pctDentro*100));
  s4 = replaceRunsSeq(s4, '0,0', [fmtNumBR(val.pctPend*100), fmtNumBR(val.pctFora*100)]);
  s4 = replaceRun(s4, '4,0', fmtNumBR(val.pctDiv*100));
  zip.file('ppt/slides/slide4.xml', s4);
  await fillValidacaoSlide(zip, 4, val.tabela.map(r=>[r.cnpj, r.nome, r.status]));

  // ---- Slides 6 e 10 — Frequência (Vendas / Estoque) ----
  log('Editando slides de Frequência...');
  async function fillFreq(slideNum, f, nome){
    let s = await zip.file('ppt/slides/slide'+slideNum+'.xml').async('string');
    s = replaceRun(s, 'Agosto 2026', mesNome);
    const t = f.total || 1;
    // ordem dos runs "%": Atualizados, Em atenção, Desatualizados
    s = replaceRunsSeq(s, '%', [
      fmtPctBR(f.atualizados/t), fmtPctBR(f.atencao/t), fmtPctBR(f.ofensores/t),
    ]);
    // ordem dos "777": Atualizados, Em atenção, Desatualizados
    s = replaceRunsSeq(s, '777', [String(f.atualizados), String(f.atencao), String(f.ofensores)]);
    zip.file('ppt/slides/slide'+slideNum+'.xml', s);
    log('  '+nome+': '+f.atualizados+' / '+f.atencao+' / '+f.ofensores+' (total '+f.total+')', 'ok');
  }
  await fillFreq(6, fv, 'Frequência de Vendas');
  await fillFreq(10, fe, 'Frequência de Estoque');

  // ---- Ofensores (Sellout: 7-9 | Estoque: 11-13) ----
  log('Editando slides de Ofensores...');
  const pctFreqOf = d => {
    const totalDias = Number(d.considerados) || 0;
    const diasSemDados = Number(d.ausentes) || 0;
    
    return totalDias > 0 ? (diasSemDados / totalDias) : 0;
  };
  const mapOf = d => [d.cnpj, d.nome, String(d.ausentes), fmtNumBR(pctFreqOf(d)*100)+'%'];
  const dataMes = (xml) => replaceRun(xml, 'Agosto 2026', mesNome);
  await fillTableAcrossSlides(zip, [7,8,9], fv.ofensoresList.map(mapOf), 'Ofensores Sellout', dataMes);
  await fillTableAcrossSlides(zip, [11,12,13], fe.ofensoresList.map(mapOf), 'Ofensores Estoque', dataMes);

  // ---- Slide 14 — Tratamento de Inconsistências ----
  log('Editando Tratamento de Inconsistências...');
  let s14 = await zip.file('ppt/slides/slide14.xml').async('string');
  const s14Vals = [
    String(as.distribuidores), String(as.impactados), String(as.associados),
    String(as.naoIdent), String(as.rejeitados),
  ];
  const s14Originais = ['8','31','31','1','0'];
  let cursor14 = 0;
  s14 = s14.replace(/<a:t>([^<]*)<\/a:t>/g, (m, txt) => {
    if (cursor14 < s14Originais.length && txt === s14Originais[cursor14]) {
      const rep = '<a:t>'+xmlEscape(s14Vals[cursor14])+'</a:t>';
      cursor14++;
      return rep;
    }
    return m;
  });
  zip.file('ppt/slides/slide14.xml', s14);

  log('Compactando arquivo...');
  const out = await zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    compression: 'DEFLATE',
  });
  log('✓ Apresentação gerada.', 'ok');
  return out as Blob;
}
