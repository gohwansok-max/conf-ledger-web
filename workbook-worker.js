/* Offline workbook I/O. Formula caches are observations, never executed. */
(function (root, factory) {
  'use strict';
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./vendor/xlsx.full.min.js'), (() => { try { return require('./ledger-core.js'); } catch (_) { return null; } })());
  else {
    if (!root.XLSX) root.importScripts('vendor/xlsx.full.min.js');
    const api = factory(root.XLSX, root.ConfLedger || root.LedgerCore);
    root.onmessage = async function (event) {
      const m = event.data || {};
      try {
        let result;
        if (m.type === 'inspect') result = api.inspect(m.buffer);
        else if (m.type === 'import') result = api.importRows(m.buffer, m.filename, m.sheet);
        else if (m.type === 'export') result = api.exportWorkbook(m.rows, m.masters, m.originalBuffer);
        else throw Error('지원하지 않는 작업입니다.');
        root.postMessage({id:m.id,ok:true,result}, m.type === 'export' ? [result.buffer] : []);
      } catch (error) { root.postMessage({id:m.id,ok:false,error:String(error.message || error)}); }
    };
  }
})(typeof globalThis !== 'undefined' ? globalThis : this, function (XLSX, ledger) {
  'use strict';
  const FALLBACK_COLUMNS = ['원본시트','스토어','월','주','발주일','주문일','유,무상','구매처','상세구매처','제품명','상세상품','수량','수량(알)','무상출고','정산가','판매가','배송비','단가','수수료율','행사','비고','벌크','띠지','구매자명','구매자연락처','수취인명','수취인 연락처','우편번호','주소','배송메세지','증빙 여부','택배사','송장번호','주문고유번호','품목번호'];
  const COLUMNS = (ledger && ledger.COLUMNS || FALLBACK_COLUMNS).map(c => typeof c === 'string' ? c : c.key || c.name || c.label);
  const SAMPLE = ['주문번호','상품명','수량','원본가격','수취인명','연락처','우편번호','주소'];
  const META = '__CONF_META';
  const META_HEADERS = ['exportRow','identifier','fingerprint','id','status','revision','sourceFile','sourceSheet','sourceRow','updatedAt','notes','confirmed'];
  const LIMIT = 50000;
  const blank = v => v === null || v === undefined || String(v).trim() === '';
  const clean = v => String(v == null ? '' : v).replace(/[\s\r\n]+/g,'').trim();
  const aliases = {상품명:'제품명',옵션명:'상세상품',상품옵션:'상세상품',옵션:'상세상품',원본가격:'판매가',수취인명:'수취인명',수령인:'수취인명',연락처:'수취인 연락처',수취인연락처:'수취인 연락처',주문번호:'주문고유번호',배송주소:'주소',주문일자:'주문일',주문날짜:'주문일'};
  const dateFields = /일자|일시|날짜|주문일|발주일|배송일|출고일|접수일/;
  const textFields = /전화|연락처|우편|번호|코드|아이디|ID$/i;
  let serial = 0;
  const uuid = () => typeof globalThis.crypto?.randomUUID === 'function' ? globalThis.crypto.randomUUID() : 'conf-' + Date.now().toString(36) + '-' + (++serial).toString(36);
  const timestamp = () => new Date().toISOString();
  function bytes(buffer) {
    if (buffer instanceof ArrayBuffer) return new Uint8Array(buffer);
    if (ArrayBuffer.isView(buffer)) return new Uint8Array(buffer.buffer,buffer.byteOffset,buffer.byteLength);
    throw Error('엑셀 파일 버퍼가 필요합니다.');
  }
  function read(buffer, options) {
    const b=bytes(buffer);if(b.byteLength>50*1024*1024)throw Error('50MB 이하 파일만 지원합니다.');
    const wb=XLSX.read(b,{type:'array',cellFormula:true,cellNF:true,cellText:true,cellDates:false,sheetStubs:true,...options});
    // SheetJS 0.18.5 synthesizes v=0 for formula cells without a cache.
    // Inspect the raw OOXML selected sheet, never accept that synthesized zero.
    if(!options?.skipCacheRepair && wb.Sheets && b[0]===80 && b[1]===75) repairMissingCaches(b,wb);
    return wb;
  }
  function xmlAttrs(tag) {const a={};for(const m of tag.matchAll(/([\w:.-]+)\s*=\s*["']([^"']*)["']/g))a[m[1]]=m[2].replace(/&quot;/g,'"').replace(/&apos;/g,"'").replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&amp;/g,'&');return a;}
  function repairMissingCaches(b,wb) {
    const zip=XLSX.CFB.read(b,{type:'array'});
    const text=p=>{const i=zip.FullPaths.findIndex(n=>n==='Root Entry/'+p);if(i<0)return '';const content=zip.FileIndex[i].content;return typeof TextDecoder==='function'?new TextDecoder('utf-8').decode(content):decodeURIComponent(Array.from(content,n=>'%'+n.toString(16).padStart(2,'0')).join(''));};
    const rels={};for(const m of text('xl/_rels/workbook.xml.rels').matchAll(/<Relationship\b[^>]*\/?\s*>/g)){const a=xmlAttrs(m[0]);if(a.Id&&a.Target)rels[a.Id]=a.Target;}
    for(const m of text('xl/workbook.xml').matchAll(/<sheet\b[^>]*\/?\s*>/g)) {
      const a=xmlAttrs(m[0]),ws=wb.Sheets[a.name],target=rels[a['r:id']];if(!ws||!target)continue;
      const components=(target.startsWith('/')?target.slice(1):'xl/'+target).split('/'),normalized=[];for(const c of components){if(c==='..')normalized.pop();else if(c&&c!=='.')normalized.push(c);}
      const xml=text(normalized.join('/'));
      for(const cm of xml.matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>/g)) {
        const attr=xmlAttrs(cm[1]),body=cm[2];if(!/<f(?:\s|>|\/)/.test(body))continue;
        if(!/<v(?:\s[^>]*)?>\s*[^<\s][^<]*<\/v>/.test(body)) {const cell=ws[attr.r] || (ws[attr.r]={t:'n',f:''});delete cell.v;cell.__missingCache=true;}
      }
    }
  }
  function requireColumns() { if(COLUMNS.length!==35 || COLUMNS.some(blank)) throw Error('발주DB 35열 계약을 확인할 수 없습니다.'); }
  function canonical(h, index, headers) {
    if(h==='비고' && headers.slice(0,index).includes('비고')) return '품목번호';
    const exact=COLUMNS.find(c=>clean(c)===h);
    if(exact) return exact;
    const target = aliases[h];
    if(target && COLUMNS.includes(target)) return target;
    // Prefer the exact original column name over a guessed alias.
    const candidates = {판매가:['원본가격','판매금액','금액','가격'],수취인:['수취인명','수령인','받는분'],수취인연락처:['연락처','수취인연락처1','수취인전화번호','전화번호'],주문일:['주문일자','발주일','주문일시'],주소:['수취인주소','배송주소']};
    for(const c of candidates[target || h] || []) if(COLUMNS.includes(c)) return c;
    return target || h;
  }
  function headerInfo(ws, strict = true) {
    if(!ws || !ws['!ref']) return null;
    const range = XLSX.utils.decode_range(ws['!ref']);
    if(range.e.r>LIMIT || range.e.c>200) {if(strict)throw Error('시트 크기가 허용 범위를 초과했습니다.');return null;}
    for(let r=range.s.r;r<=Math.min(range.e.r,range.s.r+29);r++) {
      const headers=[];
      for(let c=0;c<=range.e.c;c++) {const cell=ws[XLSX.utils.encode_cell({r,c})];headers.push(clean(cell && cell.t!=='e' ? cell.v : ''));}
      const original = ['제품명','상세상품','수량'].every(h=>headers.includes(h));
      const sample = SAMPLE.every(h=>headers.includes(h));
      if(original || sample) return {r,range,headers,sample,keys:headers.map((h,i)=>canonical(h,i,headers))};
    }
    return null;
  }
  function getData(row) { return row && (row.values || row.data || row.fields) || row || {}; }
  function rowValues(row) { const data=getData(row);return COLUMNS.map((h,i)=>Array.isArray(data)? data[i]??null : data[h]??null); }
  const fingerprint = values => JSON.stringify(values.map(v => blank(v) ? null : v));
  function identifier(values) { return ['주문고유번호','품목번호'].map(h=>String(values[COLUMNS.indexOf(h)]??'')).join('|'); }
  function rowMeta(row) { return row.meta || row; }
  function realRow(values, info) {
    const product=values['제품명'] ?? values['상품명'], option=values['상세상품'];
    if(blank(product) && blank(option)) return false;
    if(/^(총합계|총계|합계|소계|행레이블|GrandTotal)$/i.test(clean(product)) || /^(총합계|총계|합계|소계)$/i.test(clean(option))) return false;
    const keys=Object.keys(values);
    const hasDate=keys.some(k=>dateFields.test(k)&&!blank(values[k]));
    const hasBuyer=keys.some(k=>/주문자|구매자|수취인|수령인|받는분|거래처/.test(k)&&!blank(values[k]));
    return hasDate || hasBuyer;
  }
  function normalize(cell, field, notes, date1904, report) {
    if(!cell) return null;
    if(cell.t==='e' || cell.__missingCache || (cell.f && (cell.v===undefined || cell.v===null))) {
      if(cell.f || cell.__missingCache) report.formulaMissing++;
      notes.push({field,code:cell.t==='e'?'CELL_ERROR':'FORMULA_CACHE_MISSING',message:cell.t==='e'?'셀 오류: 값 확인 필요':'수식 캐시 값 없음: Excel 재계산 필요'});
      return null;
    }
    let value=cell.v;
    if(value===undefined || value===null) return null;
    if(dateFields.test(field)) {
      let parts;
      if(typeof value==='number') parts=XLSX.SSF.parse_date_code(value,{date1904});
      if(value instanceof Date) parts={y:value.getFullYear(),m:value.getMonth()+1,d:value.getDate(),H:value.getHours(),M:value.getMinutes(),S:value.getSeconds()};
      if(parts) return `${String(parts.y).padStart(4,'0')}-${String(parts.m).padStart(2,'0')}-${String(parts.d).padStart(2,'0')}` + ((parts.H || parts.M || parts.S || /일시/.test(field)) ? ` ${String(parts.H||0).padStart(2,'0')}:${String(parts.M||0).padStart(2,'0')}:${String(Math.floor(parts.S||0)).padStart(2,'0')}` : '');
      if(typeof value==='string') { const m=value.trim().match(/^(\d{4})[-./년]\s*(\d{1,2})[-./월]\s*(\d{1,2})(?:일)?(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?(?:Z|[+-]\d{2}:?\d{2})?$/); if(m) { const y=+m[1],mo=+m[2],d=+m[3], dt=new Date(y,mo-1,d);if(dt.getFullYear()===y&&dt.getMonth()===mo-1&&dt.getDate()===d) return `${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}`+(m[4]?` ${m[4].padStart(2,'0')}:${m[5]}:${m[6]||'00'}`:''); } }
    }
    if(textFields.test(field)) {
      if(/우편/.test(field) && typeof value==='number' && Number.isInteger(value) && value>=0 && value<100000) { if(value<10000) notes.push({field,code:'ZIPCODE_PADDED',message:'숫자 우편번호를 5자리로 보정'}); return String(value).padStart(5,'0'); }
      if(typeof value==='number' && cell.z && /^0+$/.test(cell.z)) return XLSX.SSF.format(cell.z,value);
      return String(value);
    }
    return value;
  }
  function extract(wb, sheet, filename, withMeta) {
    const ws=wb.Sheets[sheet],info=headerInfo(ws);
    if(!info) throw Error('지원하는 주문 또는 발주DB 열이 없는 시트입니다.');
    const report={read:0,skipped:0,formulaMissing:0,sourceSheet:sheet},rows=[];
    const date1904=!!wb.Workbook?.WBProps?.date1904;
    const metas=new Map();
    if(withMeta && wb.Sheets[META]) for(const m of XLSX.utils.sheet_to_json(wb.Sheets[META],{defval:null})) metas.set(Number(m.exportRow),m);
    for(let r=info.r+1;r<=info.range.e.r;r++) {
      const values=Object.fromEntries(COLUMNS.map(h=>[h,null])),notes=[];let present=false;
      for(let c=0;c<info.headers.length;c++) {const h=info.keys[c];if(!h||!COLUMNS.includes(h))continue;const cell=ws[XLSX.utils.encode_cell({r,c})];if(cell && (!blank(cell.v)||cell.f||cell.t==='e'))present=true;values[h]=normalize(cell,h,notes,date1904,report);}
      if(!present) continue;
      if(!realRow(values,info)) {report.skipped++;continue;}
      const row={id:uuid(),values,meta:{sourceFile:String(filename||''),sourceSheet:sheet,sourceRow:r+1,status:'접수',revision:1,updatedAt:timestamp(),confirmed:false,notes}};
      const meta=row.meta;
      const m=metas.get(r+1),v=rowValues(row);
      if(m && m.id && String(m.identifier??'')===identifier(v) && m.fingerprint===fingerprint(v)) {
        row.id=String(m.id); meta.status=String(m.status||'접수');meta.confirmed=m.confirmed===true||m.confirmed==='true';meta.revision=Number.isSafeInteger(Number(m.revision))&&Number(m.revision)>=0?Number(m.revision):1;
        meta.sourceFile=String(m.sourceFile||meta.sourceFile);meta.sourceSheet=String(m.sourceSheet||meta.sourceSheet);meta.sourceRow=Number(m.sourceRow)||meta.sourceRow;meta.updatedAt=String(m.updatedAt||meta.updatedAt);
        try {const previous=JSON.parse(m.notes||'[]');if(Array.isArray(previous))meta.notes=previous.concat(notes);}catch(_){}
      }
      if(blank(values['원본시트']))values['원본시트']=meta.sourceSheet||sheet;
      rows.push(row);report.read++;
    }
    return {rows,report};
  }
  function inspect(buffer) {
    requireColumns();const preview=read(buffer,{sheetRows:30,skipCacheRepair:true});const sheets=[];
    const names=preview.SheetNames.filter(name=>name!==META&&headerInfo(preview.Sheets[name],false));
    const wb=names.length?read(buffer,{sheets:names}):preview;
    for(let i=0;i<preview.SheetNames.length;i++){const name=preview.SheetNames[i];if(name===META)continue;const info=names.includes(name)?headerInfo(wb.Sheets[name],false):null;const count=info?extract(wb,name,'',false).rows.length:0;sheets.push({name,rows:count,compatible:!!info,hidden:!!preview.Workbook?.Sheets?.[i]?.Hidden});}
    const supported=sheets.filter(s=>s.compatible);
    const preferred=supported.find(s=>s.name==='발주DB')||supported.find(s=>!s.hidden && /온라인|online/i.test(s.name))||supported.find(s=>!s.hidden)||supported[0];
    return {sheets,defaultSheet:preferred?.name||null};
  }
  function importRows(buffer,filename,sheet) {
    requireColumns();if(!sheet)sheet=inspect(buffer).defaultSheet;if(!sheet)throw Error('지원하는 주문 시트가 없습니다.');
    const names=read(buffer,{bookSheets:true}).SheetNames;if(!names.includes(sheet)||sheet===META)throw Error('선택한 시트를 찾을 수 없습니다.');
    const wb=read(buffer,{sheets:[sheet,META,'CONF_상품기준','상품마스터']});const result=extract(wb,sheet,filename,true),masters=[];
    const masterSheet=wb.Sheets['CONF_상품기준']||wb.Sheets['상품마스터'];
    if(masterSheet) for(const m of XLSX.utils.sheet_to_json(masterSheet,{defval:null})) {const name=[m.name,m['상세상품'],m['제품명'],m['상품명']].find(v=>!blank(v));if(!blank(name))masters.push({name:String(name),pack:m.pack??m['입수']??null,status:m.status??m['상태']??'사용'});}
    return {...result,masters};
  }
  function metaRows(rows) {return [META_HEADERS,...rows.map((row,i)=>{const v=rowValues(row),m=rowMeta(row);return [i+2,identifier(v),fingerprint(v),row.id||uuid(),m.status||'접수',m.revision??1,m.sourceFile||'',m.sourceSheet||'',m.sourceRow??null,m.updatedAt||timestamp(),JSON.stringify(m.notes||[]),m.confirmed===true];})];}
  function standard(rows,masters) {
    const wb=XLSX.utils.book_new();const add=(name,aoa)=>XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(aoa),name);
    add('발주DB',[COLUMNS,...rows.map(rowValues)]);
    add('상품마스터',[['제품명','입수','상태'],...(masters||[]).map(m=>[m.name,m.pack??null,m.status||'사용'])]);
    add('처리현황',[['ID','상태','수정차수','원본파일','원본시트','원본행'],...rows.map(r=>{const m=rowMeta(r);return [r.id||'',m.status||'접수',m.revision??1,m.sourceFile||'',m.sourceSheet||'',m.sourceRow??null];})]);
    add(META,metaRows(rows));wb.Workbook={Sheets:wb.SheetNames.map(name=>({name,Hidden:name===META?2:0})),CalcPr:{fullCalcOnLoad:true,forceFullCalc:true}};
    wb.Sheets['발주DB']['!cols']=COLUMNS.map(h=>({wch:/주소|상품|비고/.test(h)?28:16}));
    wb.Sheets['발주DB']['!autofilter']={ref:wb.Sheets['발주DB']['!ref']};
    return XLSX.write(wb,{bookType:'xlsx',type:'array'});
  }
  function exportWorkbook(rows,masters,originalBuffer) {
    requireColumns();if(!Array.isArray(rows))throw Error('원장 행 목록이 필요합니다.');if(rows.length>LIMIT)throw Error('5만 행 이하만 내보낼 수 있습니다.');
    const normalized=rows.map(row=>({...row,id:row.id||uuid()}));
    const result={buffer:standard(normalized,masters),filename:'발주통합_원장.xlsx',mode:'standard',note:'35열 표준 원장입니다. 수식을 실행하지 않으며 변경이력은 추가하지 않았습니다.'};
    if(originalBuffer) result.note+=' 원본 보존 내보내기는 지원하지 않으므로 기존 16개 시트·차트·스타일 보존을 주장하지 않습니다. 원본 파일은 변경하지 않았습니다. 원본은 별도로 보관하세요.';
    return result;
  }
  return {COLUMNS,inspect,importRows,exportWorkbook};
});
