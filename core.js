// Shared, dependency-free data rules. All charts and exports use these functions.
export const norm = v => String(v ?? '').normalize('NFKC').toLowerCase().replace(/[^a-z0-9]/g, '');
export const codeKey = v => String(v ?? '').trim().toUpperCase().replace(/\s+/g, '');
export const clean = v => /^(#(N\/A|REF!|VALUE!|DIV\/0!|NAME\?|NUM!|SPILL!)|n\/a|null|undefined)$/i.test(String(v ?? '').trim()) ? '' : String(v ?? '').trim();
export function number(v) {
  if (v === '' || v == null || typeof v === 'boolean') return null;
  if (typeof v === 'number') return Number.isFinite(v) ? v : null;
  let s = String(v).trim(); if (!s || /^[-–—]$/.test(s)) return null;
  const negative = /^\(.*\)$/.test(s); const pct = s.endsWith('%');
  s = s.replace(/[৳,%\s()]/g, '').replace(/^(BDT|Tk\.?)/i, '');
  if (!/^[-+]?\d*\.?\d+(e[-+]?\d+)?$/i.test(s)) return null;
  const n = Number(s) * (negative ? -1 : 1) / (pct ? 100 : 1);
  return Number.isFinite(n) ? n : null;
}
const ALIASES = {
  code: ['code','outlet code','store code','outlet id','store id','branch code','outlet no','shop code'],
  name: ['outlet name','store name','branch name','shop name','outlet'],
  leader: ['leader','RHO','regional head','regional head name','region head','regional manager'],
  zonal: ['zonal','zone head','zonal name','zonal manager','zone manager','zonal in charge'],
  sales: ['POS & MFS Sales','POS MFS Sales','card sales','credit card sales','payment sales','POS sales'],
  extra: ['extra amount','total extra amount','credit card extra amount','total credit card extra amount','excess charge','extra charge','total extra charge','additional charge','total opportunity loss','opportunity loss'],
  projection: ['month end projection','month end projected amount','projected amount','projection','monthly projection'],
  target: ['target','target amount','saving target','savings target'],
  saving: ['save amount','saving amount','savings amount','saving','savings'],
  incentive: ['incentive amount','incentive'],
  serial: ['SL','SL No','serial','serial no'],
  format: ['format','outlet format','store format','ownership','own fr'],
  division: ['division'], district: ['district'], area: ['area'],
  location: ['location type','location category'], status: ['status','outlet status']
};
const BANKS = { bkash:'bKash',brac:'BRAC',bracbank:'BRAC',city:'City',citybank:'City',dbbl:'DBBL',dutchbanglabank:'DBBL',ebl:'EBL',easternbank:'EBL',mtb:'MTB',mutualtrustbank:'MTB',pbl:'PBL',primebank:'PBL',nagad:'Nagad',upay:'Upay',rocket:'Rocket',ucb:'UCB',ncc:'NCC',ific:'IFIC',tbl:'TBL',southeastbank:'Southeast Bank',islamibank:'Islami Bank' };
export const ROLES = ['extra','bank','sales','projection','target','saving','incentive','code','name','leader','zonal','format','division','district','area','location','status','serial','other'];
export function headerRole(header, overrides = {}) {
  if (ROLES.includes(overrides[header])) return overrides[header];
  const n = norm(header);
  for (const [r, a] of Object.entries(ALIASES)) if (a.some(v => norm(v) === n)) return r;
  if (BANKS[n]) return 'bank';
  if (/(extraamount|extracharge|excesscharge|additionalcharge)/.test(n) && !/^total/.test(n)) return 'bank';
  for (const b of Object.keys(BANKS)) if (n.startsWith(b) && /^(extra|charge|cost|amount|fee)/.test(n.slice(b.length))) return 'bank';
  return 'other';
}
export function inferColumns(headers, lines, overrides = {}, channelIndexes = []) {
  const channelSet=new Set(channelIndexes);
  const cols = headers.map((label, i) => {
    const samples = lines.map(r => r[i]).filter(v => v !== '' && v != null).slice(0, 200);
    const numeric = samples.length > 0 && samples.filter(v => number(v) != null).length / samples.length > .85;
    const role=channelSet.has(i)&&!Object.hasOwn(overrides,String(label||'').trim())?'bank':headerRole(label, overrides);
    return { key:String(label || '').trim(), index:i, label:String(label || '').trim(), role, numeric };
  }).filter(c => c.key);
  // Payment-channel columns form a contiguous block before Total Extra Amount.
  // Known and newly introduced numeric channel headings are both accepted.
  const total = cols.find(c => c.role === 'extra');
  const bankPositions = cols.filter(c => c.role === 'bank').map(c => c.index);
  const code = cols.find(c => c.role === 'code');
  if (total && (bankPositions.length || code)) {
    const first = bankPositions.length ? Math.min(...bankPositions) : code.index;
    for (const c of cols) if (c.role === 'other' && c.numeric && c.index > first && c.index < total.index && !Object.hasOwn(overrides,c.key) && !/sales|rate|percent|%|count|transaction|target|projection|amountpaid/i.test(c.key)) c.role = 'bank';
  }
  const used = new Set();
  for (const c of cols) { if (used.has(c.key)) throw new Error(`Duplicate column heading: ${c.key}. Give each column a distinct name.`); used.add(c.key); }
  return cols;
}
function stackHeaderRows(upper, lower, overrides = {}) {
  const width=Math.max(upper?.length||0,lower?.length||0);const headers=[];
  const channelStart=upper.findIndex(v=>/^(payment)?channels?$/.test(norm(v)));
  const totalIndex=channelStart>=0?upper.findIndex((v,i)=>i>channelStart&&/^(grand)?total/.test(norm(v))):-1;
  const channelIndexes=channelStart>=0&&totalIndex>channelStart?Array.from({length:totalIndex-channelStart-1},(_,i)=>channelStart+i+1):[];
  for(let i=0;i<width;i++) {
    const top=clean(upper?.[i]);const bottom=clean(lower?.[i]);
    if(channelIndexes.includes(i)){headers[i]=top||bottom;continue;}
    if(!top){headers[i]=bottom;continue;}if(!bottom){headers[i]=top;continue;}
    const topRole=headerRole(top,overrides);const bottomRole=headerRole(bottom,overrides);
    if(bottomRole!=='other'&&bottomRole!=='extra'){headers[i]=bottom;continue;}
    if(topRole==='bank'){headers[i]=top;continue;}
    if(bottomRole==='extra'&&norm(top)==='total'){headers[i]=bottom;continue;}
    headers[i]=topRole!=='other'?top:`${top} ${bottom}`;
  }
  return {headers,channelIndexes};
}
export function detectTable(sheets, kind, overrides = {}) {
  const candidates = [], diagnostics = [], seen = new Set();
  const consider=(s,h,heads,preambleEnd,stacked=false,channelIndexes=[])=>{
    const roles = heads.map(v => headerRole(v, overrides));
    if (!roles.includes('code')) return;
    const isZone = roles.includes('leader') && roles.includes('zonal') && (roles.includes('format') || roles.includes('division') || roles.includes('district') || roles.includes('location'));
    if (kind === 'zone' && !isZone) return;
    if (kind === 'credit' && !roles.includes('extra') && !roles.includes('bank')) return;
    if (kind === 'credit' && isZone && !roles.includes('extra')) return;
    let cols;try{cols = inferColumns(heads, s.grid.slice(h+1), overrides, channelIndexes);}catch(e){diagnostics.push(`${s.name}: ${e.message}`);return;}
    const signature=`${s.name}\u0000${h}\u0000${cols.map(c=>`${c.index}:${c.role}:${norm(c.key)}`).join('|')}`;
    if(seen.has(signature))return;seen.add(signature);
    const score = (kind === 'credit' ? (roles.includes('extra') ? 100 : 0) + cols.filter(c=>c.role==='bank').length * 4 : roles.filter(r=>r!=='other').length) + Math.min(s.grid.length / 10000, .9);
    candidates.push({ sheetName:s.name, headerRow:h+1, headers:heads, columns:cols, lines:s.grid.slice(h+1), preamble:s.grid.slice(0,preambleEnd), score, stacked });
  };
  for (const s of sheets) for (let h = 0; h < Math.min(s.grid.length, 40); h++) {
    const heads=s.grid[h]||[];consider(s,h,heads,h,false);
    if(h>0&&heads.some(v=>headerRole(v,overrides)==='code')){const stacked=stackHeaderRows(s.grid[h-1]||[],heads,overrides);consider(s,h,stacked.headers,h-1,true,stacked.channelIndexes);}
  }
  candidates.sort((a,b)=>b.score-a.score);
  if (!candidates.length) throw new Error((kind === 'zone' ? 'No Zone Distribution table found. Required: outlet CODE, Leader, Zonal and a location or format column.' : 'No credit card table found. Required: Outlet Code and Extra Amount or payment-channel columns. Use Column rules for renamed headings.') + (diagnostics.length?' '+diagnostics[0]:''));
  if(candidates[1] && Math.abs(candidates[0].score-candidates[1].score)<.00001) throw new Error(`Two equally matching ${kind} tables found. Keep one current source table.`);
  return candidates[0];
}
function get(table, row, role) { const c=table.columns.find(c=>c.role===role); return c ? row[c.index] ?? '' : ''; }
function hash(s) { let n=2166136261; for(const c of s) n=Math.imul(n^c.charCodeAt(0),16777619); return (n>>>0).toString(36); }
export function monthEndProjection(extra, meta = {}) {
  const amount=number(extra);const tillDays=number(meta.tillDays??meta.elapsedDays);const monthDays=number(meta.monthDays);
  return amount!==null&&tillDays>0&&monthDays>0?amount/tillDays*monthDays:null;
}
const MONTH_LABELS=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
function modifiedPeriod(value) {
  const date=new Date(value);if(!Number.isFinite(date.getTime()))return null;
  let day,month,year;
  try {
    const parts=Object.fromEntries(new Intl.DateTimeFormat('en-US',{timeZone:'Asia/Dhaka',day:'numeric',month:'numeric',year:'numeric'}).formatToParts(date).map(p=>[p.type,p.value]));
    day=+parts.day;month=+parts.month;year=+parts.year;
  } catch {day=date.getUTCDate();month=date.getUTCMonth()+1;year=date.getUTCFullYear();}
  const monthDays=new Date(Date.UTC(year,month,0)).getUTCDate();
  return day>=1&&day<=monthDays?{day,monthDays,period:`1–${day} ${MONTH_LABELS[month-1]} ${year}`} : null;
}
function metadata(table, source) {
  let period=''; const text=[...table.preamble,...table.lines.slice(-10)].flat().filter(v=>typeof v==='string').join(' ').replace(/\s+/g,' ');
  const m=text.match(/From\s*(\d{1,2})\s*to\s*(\d{1,2})\s*([A-Za-z]+)[_\s,-]*(20\d{2})/i);
  let elapsedDays=null, tillDays=null, monthDays=null;
  if(m) {
    const month=new Date(`${m[3]} 1, ${m[4]}`).getMonth();const fromDay=+m[1];const toDay=+m[2];
    if(Number.isFinite(month)) {
      const daysInMonth=new Date(+m[4],month+1,0).getDate();
      if(fromDay>=1&&toDay>=fromDay&&toDay<=daysInMonth){tillDays=toDay;elapsedDays=toDay;monthDays=daysInMonth;period=`${m[1]}–${m[2]} ${m[3]} ${m[4]}`;}
    }
  }
  if(!period) {
    const f=text.match(/Transaction\s+Date\s+is\s+on\s+or\s+after\s+(?:[A-Za-z]+,\s*)?([A-Za-z]+)\s+(\d{1,2}),\s*(20\d{2})\s+and\s+is\s+before\s+(?:[A-Za-z]+,\s*)?([A-Za-z]+)\s+(\d{1,2}),\s*(20\d{2})/i);
    if(f) {
      const startMonth=MONTH_LABELS.findIndex(v=>f[1].toLowerCase().startsWith(v.toLowerCase()));
      const endMonth=MONTH_LABELS.findIndex(v=>f[4].toLowerCase().startsWith(v.toLowerCase()));
      if(startMonth>=0&&endMonth>=0) {
        const start=new Date(Date.UTC(+f[3],startMonth,+f[2]));const endExclusive=new Date(Date.UTC(+f[6],endMonth,+f[5]));const end=new Date(endExclusive.getTime()-86400000);
        if(end>=start){tillDays=end.getUTCDate();elapsedDays=tillDays;monthDays=new Date(Date.UTC(end.getUTCFullYear(),end.getUTCMonth()+1,0)).getUTCDate();period=start.getUTCMonth()===end.getUTCMonth()&&start.getUTCFullYear()===end.getUTCFullYear()?`${start.getUTCDate()}–${end.getUTCDate()} ${MONTH_LABELS[end.getUTCMonth()]} ${end.getUTCFullYear()}`:`${start.getUTCDate()} ${MONTH_LABELS[start.getUTCMonth()]}–${end.getUTCDate()} ${MONTH_LABELS[end.getUTCMonth()]} ${end.getUTCFullYear()}`;}
      }
    }
  }
  if(!period) {
    const inferred=modifiedPeriod(source.modifiedTime);
    if(inferred){tillDays=inferred.day;elapsedDays=inferred.day;monthDays=inferred.monthDays;period=inferred.period;}
  }
  let targetTotal=null; const tc=table.columns.find(c=>c.role==='target');
  if(tc) for(const row of table.preamble) {const n=number(row[tc.index]);if(n!==null&&n>1)targetTotal=n;}
  return {...source,sheetName:table.sheetName,headerRow:table.headerRow,period,elapsedDays,tillDays,monthDays,targetTotal};
}
export function parseCredit(sheets, source = {}, overrides = {}) {
  const t=detectTable(sheets,'credit',overrides); const records=[]; const banks=t.columns.filter(c=>c.role==='bank').map(c=>({key:c.key,label:BANKS[norm(c.key)]||c.key}));
  const meta=metadata(t,source);
  const warnings=[];let invalidNumbers=0, totalMismatch=0, projectedFromRule=0, projectionUnavailable=0;
  for(let i=0;i<t.lines.length;i++) {
    const row=t.lines[i]||[]; const raw=Object.fromEntries(t.columns.map(c=>[c.key,row[c.index]??'']));
    const code=clean(get(t,row,'code')); const name=clean(get(t,row,'name')); const serial=clean(get(t,row,'serial'));
    if (/^(grand\s*total|total|sub\s*total)$/i.test(code||name||serial) || /^outlet\s*code$/i.test(code) || /^applied\s*filters?:/i.test(code||name||serial)) continue;
    if(t.stacked&&i===0&&!code&&!name&&!serial)continue;
    const numericCols=t.columns.filter(c=>['bank','extra','sales','projection','target','saving','incentive'].includes(c.role));
    if(!code&&!name&&!numericCols.some(c=>number(row[c.index])!==null&&number(row[c.index])!==0)) continue;
    const values={};for(const b of banks) { const v=raw[b.key]; values[b.key]=number(v); if(clean(v)&&number(v)===null)invalidNumbers++; }
    const computed=banks.length?Object.values(values).reduce((a,v)=>a+(v??0),0):null;
    const reported=number(get(t,row,'extra'));const extra=reported??computed;
    if(reported!==null&&computed!==null&&Math.abs(reported-computed)>.011)totalMismatch++;
    if(extra===null) {warnings.push(`Row ${t.headerRow+i+1}: Extra Amount is unavailable.`);}
    const id=code?codeKey(code):`BLANK-${hash(JSON.stringify([name,raw]))}-${i}`;
    // Month End Projection is always Total Extra Amount ÷ till-day count in
    // the month × total days in that month, computed here rather than trusted
    // from any saved Projection column in the source file, so the rule stays
    // consistent across every workbook regardless of what it has saved.
    const projection=monthEndProjection(extra,meta);
    if(projection!==null)projectedFromRule++;
    else projectionUnavailable++;
    records.push({id,code:codeKey(code),name,sourceLeader:clean(get(t,row,'leader')),sourceZonal:clean(get(t,row,'zonal')),extra,projection,target:number(get(t,row,'target')),saving:number(get(t,row,'saving')),incentive:number(get(t,row,'incentive')),sales:number(get(t,row,'sales')),bankValues:values,raw,sourceRow:t.headerRow+i+1});
  }
  if(!records.length)throw new Error('The credit card table has no outlet data. The previous snapshot has been kept.');
  if(invalidNumbers)warnings.push(`${invalidNumbers} payment-channel cells contain non-numeric values. Blank cells are not assumed to be confirmed zero.`);
  if(totalMismatch)warnings.push(`${totalMismatch} rows have a reported Extra Amount different from their channel sum. Reported totals are used; review changed formulas or column rules.`);
  if(projectedFromRule)warnings.push(`Month End Projection for ${projectedFromRule} row${projectedFromRule===1?'':'s'} is calculated as Total Extra Amount ÷ ${meta.tillDays??meta.elapsedDays} till days × ${meta.monthDays} total days in the month. Any saved Projection value in the source file is not used.`);
  if(projectionUnavailable)warnings.push(`Month End Projection is unavailable for ${projectionUnavailable} row${projectionUnavailable===1?'':'s'} because the till-day count or total days in the month could not be read from the file.`);
  const counts=new Map();records.forEach(r=>{if(r.code)counts.set(r.code,(counts.get(r.code)||0)+1);});
  const dup=[...counts].filter(([,v])=>v>1);if(dup.length)warnings.push(`${dup.length} repeated outlet codes: all source rows are included, not silently removed.`);
  return {records,banks,columns:t.columns,meta,preamble:t.preamble,warnings,duplicateCodes:dup.map(([k])=>k)};
}
export function parseZone(sheets, source={}) {
  const t=detectTable(sheets,'zone');const rows=[];
  for(const row of t.lines) {
    const code=codeKey(clean(get(t,row,'code')));if(!code||/^(total|grandtotal|code)$/.test(norm(code)))continue;
    rows.push(Object.fromEntries(['code','name','leader','zonal','format','division','district','area','location','status'].map(r=>[r,r==='code'?code:clean(get(t,row,r))])));
  }
  if(!rows.length)throw new Error('The Zone Distribution table is empty. The previous mappings have been kept.');
  return {rows,meta:{...source,sheetName:t.sheetName,headerRow:t.headerRow}};
}
export function zoneFromSnapshot(snapshot) {
  if(!Array.isArray(snapshot?.rows)||!snapshot.rows.length)throw new Error('No published Zone Distribution rows.');
  const keys=[...new Set(snapshot.rows.flatMap(r=>Object.keys(r)))];
  return parseZone([{name:snapshot.sheetName||'Zone snapshot',grid:[keys,...snapshot.rows.map(r=>keys.map(k=>r[k]??''))]}],{fileName:snapshot.fileName,savedAt:snapshot.savedAt,signature:snapshot.fileSignature});
}
export function joinRows(credit,zone,overrides={}) {
  const index=new Map(); for(const r of zone?.rows||[]) {const key=codeKey(r.code);if(!index.has(key))index.set(key,[]);index.get(key).push(r);}
  return credit.records.map((r,idx)=>{
    const matches=index.get(r.code)||[];const distinct=[...new Set(matches.map(m=>JSON.stringify(m)))];
    const z=distinct.length===1?matches[0]:null;const manual=overrides[r.id]||{};
    const blankCode=!r.code;
    const manualLeader=clean(manual.leader),manualZonal=clean(manual.zonal);
    const resolvedLeader=manualLeader||clean(z?.leader),resolvedZonal=manualZonal||clean(z?.zonal);
    const manualComplete=Boolean(!blankCode&&manualLeader&&manualZonal);
    const codeMatchComplete=Boolean(!blankCode&&z&&resolvedLeader&&resolvedZonal);
    const mappingEligible=manualComplete||codeMatchComplete;
    const hasManualMapping=Boolean(manualLeader||manualZonal);
    let mapping=!zone?'unverified':matches.length===0?'unmatched':distinct.length>1?'conflict':(!resolvedLeader||!resolvedZonal)?'incomplete':'matched';
    if(mappingEligible&&hasManualMapping)mapping='manual';
    const leader=resolvedLeader||(!zone?clean(r.sourceLeader):'')||'Unassigned';
    const zonal=resolvedZonal||(!zone?clean(r.sourceZonal):'')||'Unassigned';
    const autoExcluded=!mappingEligible,manualExcluded=manual.excluded===true;
    return {...r,projection:monthEndProjection(r.extra,credit.meta),uid:`${r.id}:${idx}`,name:z?.name||r.name||'Unnamed outlet',leader,zonal,format:z?.format||'Unspecified',division:z?.division||'Unspecified',district:z?.district||'Unspecified',area:z?.area||'',location:z?.location||'Unspecified',outletStatus:z?.status||'Unspecified',mapping,blankCode,mappingEligible,autoExcluded,manualExcluded,excluded:autoExcluded||manualExcluded,manual};
  });
}
export function metric(row,key,selectedBanks=[]) {
  if(key==='extra'&&selectedBanks.length) return selectedBanks.reduce((s,k)=>s+(row.bankValues[k]??0),0);
  if(key.startsWith('bank:')) return row.bankValues[key.slice(5)]??null;
  if(key.startsWith('raw:'))return row.raw[key.slice(4)]??null;
  return row[key]??null;
}
export function filterRows(rows,filters={},search='',banks=[],exclude=true) {
  const q=String(search).toLowerCase().trim();
  return rows.filter(r=>{
    if(exclude&&r.excluded)return false;
    for(const [k,values] of Object.entries(filters)) if(values?.length&&!values.includes(String(r[k]??'')))return false;
    if(banks.length&&!banks.some(b=>Math.abs(r.bankValues[b]??0)>0))return false;
    return !q||[r.code,r.name,r.leader,r.zonal,r.format,r.division,r.district,...Object.values(r.raw)].join(' ').toLowerCase().includes(q);
  });
}
export function totals(rows,banks=[]) {
  const out={rows:rows.length,outlets:new Set(rows.map(r=>r.code||r.id)).size,charged:0,extra:null,projection:null,target:null,saving:null,incentive:null,sales:null};
  for(const k of ['extra','projection','target','saving','incentive','sales']){
    const nums=rows.map(r=>metric(r,k,banks)).filter(n=>typeof n==='number'&&Number.isFinite(n));out[k]=nums.length?nums.reduce((s,n)=>s+n,0):null;
    out[k+'Coverage']=nums.length;
  }
  out.charged=new Set(rows.filter(r=>(metric(r,'extra',banks)??0)>0).map(r=>r.code||r.id)).size;
  out.mapped=rows.filter(r=>['matched','manual'].includes(r.mapping)).length;
  return out;
}
export function groups(rows,key,banks=[]){
  const by=new Map();for(const r of rows){const g=String(r[key]||'Unassigned');if(!by.has(g))by.set(g,[]);by.get(g).push(r);}
  return [...by].map(([name,list])=>({name,...totals(list,banks),rows:list})).sort((a,b)=>(b.extra??-Infinity)-(a.extra??-Infinity)||a.name.localeCompare(b.name));
}
export function sortRows(rows,key,direction='desc',banks=[]) {
  const sign=direction==='asc'?1:-1;
  return [...rows].sort((a,b)=>{const x=metric(a,key,banks),y=metric(b,key,banks);if(x==null||x==='')return y==null||y===''?0:1;if(y==null||y==='')return -1;const nx=number(x),ny=number(y);if(key.startsWith('raw:')&&nx!==null&&ny!==null)return sign*(nx-ny);return sign*(typeof x==='number'&&typeof y==='number'?x-y:String(x).localeCompare(String(y),undefined,{numeric:true,sensitivity:'base'}));});
}
export function csvCell(v) {let s=String(v??'');if(/^[=+@\t\r]/.test(s)||(/^-/ .test(s)&&number(s)===null))s="'"+s;return '"'+s.replace(/"/g,'""')+'"';}
export function csv(headers,rows) {return '\uFEFF'+[headers,...rows].map(r=>r.map(csvCell).join(',')).join('\r\n');}
