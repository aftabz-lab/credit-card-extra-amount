import {readWorkbook} from './xlsx-reader.js?v=20260911-5';
import {parseCredit,parseZone,zoneFromSnapshot,isCcolOutletsSource,isIgnoredRawSource} from './core.js?v=20260918-1';
import {readCloudSnapshot,publishCloudSnapshotIfUnchanged,getPublisherSession,signInPublisher,signOutPublisher} from './supabase-sync.js';
export {getPublisherSession,signInPublisher,signOutPublisher};
export const FOLDER_ID='16HTr8nfPz4P2PMr4QB0bjgwiD110Qd-0';
export const SNAPSHOT_KEY='credit-card-extra-amount';
const CLOUD_URL='https://wstxgbzmsbosinmhhjbl.supabase.co';
const PUBLIC_KEY='sb_publishable_MCw_J7uorsKtmmokW1OpCg_Ej5DURhw';
const TABLE='dashboard_snapshots';
export async function deadline(promise,ms=20000){let timer;try{return await Promise.race([promise,new Promise((_,reject)=>{timer=setTimeout(()=>reject(new Error('Connection timed out. The last valid data is still available.')),ms);})]);}finally{clearTimeout(timer);}}
function db(){return new Promise((resolve,reject)=>{const r=indexedDB.open('credit-card-dashboard-v1',1);r.onupgradeneeded=()=>r.result.createObjectStore('data');r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error);});}
export async function localGet(key){try{const d=await db();return await new Promise((resolve,reject)=>{const t=d.transaction('data');const r=t.objectStore('data').get(key);r.onsuccess=()=>resolve(r.result??null);r.onerror=()=>reject(r.error);t.oncomplete=()=>d.close();});}catch{return null;}}
export async function localPut(key,value){const d=await db();return await new Promise((resolve,reject)=>{const t=d.transaction('data','readwrite');t.objectStore('data').put(value,key);t.oncomplete=()=>{d.close();resolve();};t.onerror=()=>{d.close();reject(new Error('This browser could not save the local changes. Export a mapping backup before closing.'));};});}
export async function readPublished(){const r=await deadline(readCloudSnapshot(SNAPSHOT_KEY));const p=r?.payload;if(!p)return null;if(p.format!=='credit-card-snapshot-v1'||!p.credit?.records?.length||!Array.isArray(p.credit.columns))throw new Error('The published credit card snapshot has an unsupported format.');return {...p,cloudUpdatedAt:r.updated_at,sourceKind:'published'};}
export async function readPublishedZone(){const r=await deadline(readCloudSnapshot('zone-distribution'));if(!r?.payload?.snapshot)return null;return zoneFromSnapshot({...r.payload.snapshot,savedAt:r.payload.snapshot.savedAt||r.updated_at});}
export async function publishedTimes(){const url=new URL(`${CLOUD_URL}/rest/v1/${TABLE}`);url.searchParams.set('select','snapshot_key,updated_at');url.searchParams.set('snapshot_key',`in.("${SNAPSHOT_KEY}","zone-distribution")`);const r=await fetch(url,{headers:{apikey:PUBLIC_KEY},cache:'no-store',signal:AbortSignal.timeout(15000)});if(!r.ok)throw new Error(`Snapshot check failed (${r.status}).`);return Object.fromEntries((await r.json()).map(r=>[r.snapshot_key,r.updated_at]));}
export async function fingerprint(payload){const bytes=new TextEncoder().encode(JSON.stringify({credit:payload.credit,zone:payload.zone,overrides:payload.overrides||{},schemaOverrides:payload.schemaOverrides||{}}));const h=await crypto.subtle.digest('SHA-256',bytes);return [...new Uint8Array(h)].map(v=>v.toString(16).padStart(2,'0')).join('');}
export async function publish(payload,baseCloudTime=null){
  if(!getPublisherSession())throw new Error('Sign in with your existing snapshot publisher account first.');
  if(!payload.credit?.records?.length||!payload.zone?.rows?.length)throw new Error('Read both Credit Card and Zone Distribution before publishing. Workbook-only leadership cannot be published as verified.');
  if(payload.credit.warnings?.some(w=>w.includes('reported Extra Amount different')))throw new Error('Resolve the Extra Amount / channel-sum mismatch using Column rules or the Excel formulas before publishing.');
  const current=await readPublished();const signature=await fingerprint(payload);
  if(current?.signature===signature)return {published:false,payload:current,reason:'unchanged'};
  if(current&&baseCloudTime&&Date.parse(current.cloudUpdatedAt)>Date.parse(baseCloudTime))throw new Error('Another publisher has saved a newer snapshot. Export your mapping backup, choose Use published data, then reapply your changes.');
  if(current&&!baseCloudTime)throw new Error('A snapshot was published on another device. Load it using Use published data before publishing this draft.');
  for(const kind of ['credit','zone']) {
    const old=Date.parse(current?.[kind]?.meta?.modifiedTime||'');const incoming=Date.parse(payload[kind]?.meta?.modifiedTime||'');
    if(Number.isFinite(old)&&Number.isFinite(incoming)&&incoming<old)throw new Error(`The ${kind} source is older than the published version. Read the latest file first.`);
  }
  const saved={format:'credit-card-snapshot-v1',schemaVersion:1,credit:payload.credit,zone:payload.zone,overrides:payload.overrides||{},schemaOverrides:payload.schemaOverrides||{},signature,generatedAt:new Date().toISOString()};
  const row=await publishCloudSnapshotIfUnchanged(SNAPSHOT_KEY,saved,current?.cloudUpdatedAt||null);
  const committed={...row.payload,cloudUpdatedAt:row.updated_at,sourceKind:'published'};
  return {published:true,payload:committed};
}
const candidate=meta=>!meta.name.startsWith('~$')&&(/\.(xlsx|xlsm|csv|tsv)$/i.test(meta.name)||meta.mimeType==='application/vnd.google-apps.spreadsheet');
const DRIVE_FOLDER_MIME='application/vnd.google-apps.folder';
// The old reference workbook is retired as a source. It is never read again,
// even as a fallback, even if it is still sitting in the shared Drive folder
// or has a newer modified time than the current "Credit Card.xlsx" file.
const isRetiredSource=meta=>/compiled[\s_-]*credit[\s_-]*card[\s_-]*extra[\s_-]*amount/i.test(meta.name);
async function listDriveTree(Drive){
  const files=[];const queue=[{id:FOLDER_ID,path:''}];const seen=new Set();
  while(queue.length){
    const folder=queue.shift();if(seen.has(folder.id))continue;seen.add(folder.id);
    const children=await deadline(Drive.listFolderFiles(folder.id));
    for(const meta of children){
      const drivePath=folder.path?`${folder.path}/${meta.name}`:meta.name;
      if(meta.mimeType===DRIVE_FOLDER_MIME)queue.push({id:meta.id,path:drivePath});
      else files.push({...meta,drivePath});
    }
  }
  return files;
}
async function download(meta){
  const Drive=window.ShwapnoDrive;
  if(meta.mimeType!=='application/vnd.google-apps.spreadsheet')return await deadline(Drive.downloadFile(meta),45000);
  const token=Drive.cachedToken()?.token;if(!token)throw new Error('Reconnect Google Drive.');
  const url=`https://www.googleapis.com/drive/v3/files/${encodeURIComponent(meta.id)}/export?mimeType=application%2Fvnd.openxmlformats-officedocument.spreadsheetml.sheet`;
  const r=await fetch(url,{headers:{Authorization:`Bearer ${token}`},signal:AbortSignal.timeout(45000)});if(!r.ok)throw new Error(`Could not export ${meta.name} (${r.status}).`);
  return new File([await r.blob()],`${meta.name}.xlsx`,{lastModified:Date.parse(meta.modifiedTime)||Date.now()});
}
const parsedCache=new Map();
export async function readDrive(schemaOverrides={},onStatus=()=>{}){
  const Drive=window.ShwapnoDrive;if(!Drive.cachedToken())throw new Error('Click Connect Google Drive to authorize this browser.');
  const listed=await listDriveTree(Drive);const all=listed.filter(candidate).filter(m=>!isRetiredSource(m)&&!isIgnoredRawSource(m)).sort((a,b)=>Date.parse(b.modifiedTime)-Date.parse(a.modifiedTime)||a.name.localeCompare(b.name));
  if(!all.length)throw new Error('No Excel or Google Sheets files were found in the configured source folder.');
  const creditFiles=all.filter(isCcolOutletsSource);const zoneFiles=all.filter(m=>!isCcolOutletsSource(m));
  if(!creditFiles.length)throw new Error('No ccol_outlets_<start-date>_to_<end-date> source file was found in the configured Drive folder or its subfolders. The retired data file is intentionally ignored; the previous snapshot has been kept.');
  // The dated CCoL export is the only Credit Card source. Zone Distribution
  // remains structure-validated independently, including in subfolders.
  let credit=null,zone=null;const errors=[];const skipped=[];
  async function tryFile(meta,{forCredit,forZone}){
    if(Number(meta.size)>30*1024*1024){skipped.push(meta.name);return;}
    onStatus(`Checking ${meta.drivePath||meta.name}`);
    const signature=Drive.remoteSignature(meta);let sheets=parsedCache.get(signature);
    const displayName=meta.drivePath||meta.name;
    try{if(!sheets){sheets=await readWorkbook(await download(meta));if(parsedCache.size>10)parsedCache.clear();parsedCache.set(signature,sheets);}}catch(e){errors.push({name:displayName,time:Date.parse(meta.modifiedTime),error:e.message});return;}
    const source={fileName:meta.name,filePath:displayName,fileId:meta.id,modifiedTime:meta.modifiedTime,signature,folderId:FOLDER_ID};
    if(forCredit&&!credit)try{credit=parseCredit(sheets,source,schemaOverrides);}catch(e){errors.push({name:displayName,time:Date.parse(meta.modifiedTime),error:e.message,kind:'credit'});}
    if(forZone&&!zone)try{zone=parseZone(sheets,source);}catch(e){errors.push({name:displayName,time:Date.parse(meta.modifiedTime),error:e.message,kind:'zone'});}
  }
  for(const meta of creditFiles){if(credit)break;await tryFile(meta,{forCredit:true,forZone:false});}
  for(const meta of zoneFiles){if(zone)break;await tryFile(meta,{forCredit:false,forZone:true});}
  if(!credit||!zone)throw new Error(`Could not validate ${!credit?'Credit Card':''}${!credit&&!zone?' and ':''}${!zone?'Zone Distribution':''}. ${errors.filter(e=>e.kind).slice(0,2).map(e=>e.name+': '+e.error).join(' ')||errors.slice(0,2).map(e=>e.name+': '+e.error).join(' ')}${skipped.length?' Large files skipped: '+skipped.join(', '):''}`);
  return {credit,zone};
}
export async function importFiles(files,schemaOverrides={}){
  let credit=null,zone=null;const notes=[];
  for(const file of files){const sheets=await readWorkbook(file);const source={fileName:file.name,modifiedTime:new Date(file.lastModified).toISOString(),signature:`${file.name}|${file.size}|${file.lastModified}`,importedLocally:true};let matched=false;
    try{const c=parseCredit(sheets,source,schemaOverrides);if(credit)throw Object.assign(new Error('Select only one Credit Card workbook per import.'),{duplicate:true});credit=c;matched=true;}catch(e){if(e.duplicate)throw e;notes.push(`${file.name}: ${e.message}`);}
    try{const z=parseZone(sheets,source);if(zone)throw Object.assign(new Error('Select only one Zone Distribution workbook per import.'),{duplicate:true});zone=z;matched=true;}catch(e){if(e.duplicate)throw e;}
    if(!matched)throw new Error(notes[notes.length-1]||`${file.name} is not a recognized source.`);
  }
  return {credit,zone};
}
