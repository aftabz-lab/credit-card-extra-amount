// XLSX/ XLSM read-only parser: cached formula values, inline/shared text, and merged headings.
// No macros or workbook formula code is executed.
const utf8=new TextDecoder();
function xml(text){const d=new DOMParser().parseFromString(text,'application/xml');if(d.getElementsByTagName('parsererror').length)throw new Error('Invalid workbook XML.');return d;}
const elements=(d,n)=>Array.from(d.getElementsByTagNameNS('*',n));
function col(s){let n=0;for(const c of s.match(/^[A-Z]+/i)?.[0]||'')n=n*26+c.toUpperCase().charCodeAt(0)-64;return n-1;}
export async function unzip(buffer) {
  const v=new DataView(buffer); const bytes=new Uint8Array(buffer);let end=-1;
  for(let i=bytes.length-22;i>=Math.max(0,bytes.length-65557);i--)if(v.getUint32(i,true)===0x06054b50){end=i;break;}
  if(end<0)throw new Error('This is not an XLSX workbook. Save older .xls files as .xlsx in Excel.');
  const entries=v.getUint16(end+10,true);let p=v.getUint32(end+16,true);const files=new Map();let total=0;
  for(let i=0;i<entries;i++) {
    if(v.getUint32(p,true)!==0x02014b50)throw new Error('Workbook ZIP directory is damaged.');
    const flags=v.getUint16(p+8,true),method=v.getUint16(p+10,true),size=v.getUint32(p+20,true),unpacked=v.getUint32(p+24,true),nameLen=v.getUint16(p+28,true),extra=v.getUint16(p+30,true),comment=v.getUint16(p+32,true),offset=v.getUint32(p+42,true);
    if(flags&1)throw new Error('Password-protected workbooks must be saved without a password for import.');
    total+=unpacked;if(total>200*1024*1024)throw new Error('Uncompressed workbook exceeds 200 MB. Remove unused sheets.');
    const name=utf8.decode(bytes.subarray(p+46,p+46+nameLen));files.set(name,{method,size,offset});p+=46+nameLen+extra+comment;
  }
  return async name=>{
    const e=files.get(name);if(!e)return null;
    const start=e.offset+30+v.getUint16(e.offset+26,true)+v.getUint16(e.offset+28,true);const compressed=bytes.subarray(start,start+e.size);
    if(e.method===0)return utf8.decode(compressed);
    if(e.method!==8)throw new Error('Unsupported workbook compression. Re-save as .xlsx.');
    let ds;try{ds=new DecompressionStream('deflate-raw');}catch{throw new Error('Use an updated Chrome, Edge, Opera, Firefox or Safari browser to import Excel files.');}
    return await new Response(new Blob([compressed]).stream().pipeThrough(ds)).text();
  };
}
export function parseCSV(text) {
  const rows=[];let row=[],s='',quoted=false;const sep=text.split('\n')[0].includes('\t')?'\t':',';
  for(let i=0;i<text.length;i++){const c=text[i];if(c==='"'){if(quoted&&text[i+1]==='"'){s+='"';i++;}else quoted=!quoted;}else if(c===sep&&!quoted){row.push(s);s='';}else if((c==='\n'||c==='\r')&&!quoted){if(c==='\r'&&text[i+1]==='\n')i++;row.push(s);rows.push(row);row=[];s='';}else s+=c;}
  if(s||row.length){row.push(s);rows.push(row);}if(rows[0]?.[0])rows[0][0]=rows[0][0].replace(/^\uFEFF/,'');return rows;
}
export async function readWorkbook(file) {
  if(file.size>30*1024*1024)throw new Error('File exceeds 30 MB. Use a workbook containing the required tables only.');
  if(/\.(csv|tsv)$/i.test(file.name))return [{name:file.name,grid:parseCSV(await file.text())}];
  const read=await unzip(await file.arrayBuffer());const wb=await read('xl/workbook.xml');if(!wb)throw new Error('Excel workbook metadata is missing.');
  const rels=xml(await read('xl/_rels/workbook.xml.rels'));const relmap=new Map(elements(rels,'Relationship').map(e=>[e.getAttribute('Id'),e.getAttribute('Target')]));
  const sharedText=await read('xl/sharedStrings.xml');const strings=sharedText?elements(xml(sharedText),'si').map(e=>elements(e,'t').map(t=>t.textContent).join('')):[];
  const sheets=[];
  for(const sheet of elements(xml(wb),'sheet')){
    const target=relmap.get(sheet.getAttribute('r:id')||sheet.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id'));if(!target)continue;
    const path=target.startsWith('/')?target.slice(1):'xl/'+target;const text=await read(path);if(!text)continue;const doc=xml(text);const grid=[];
    let nextRowIndex=0;
    for(const row of elements(doc,'row')){
      const explicitRow=Number(row.getAttribute('r'));const rowIndex=Number.isInteger(explicitRow)&&explicitRow>0?explicitRow-1:nextRowIndex;nextRowIndex=rowIndex+1;if(rowIndex>200000)throw new Error('Workbook has too many formatted rows. Remove unused rows and save again.');
      const line=[];let nextColumnIndex=0;
      for(const cell of elements(row,'c')){
        const reference=cell.getAttribute('r')||'';const i=reference?col(reference):nextColumnIndex;nextColumnIndex=i+1;if(i<0||i>500)continue;const type=cell.getAttribute('t');let val=elements(cell,'v')[0]?.textContent??'';
        if(type==='s')val=strings[Number(val)]??'';else if(type==='inlineStr')val=elements(cell,'t').map(t=>t.textContent).join('');else if(!type&&val!=='')val=Number(val);else if(type==='b')val=val==='1';
        line[i]=val;
      }
      grid[rowIndex]=line;
    }
    // Preserve row positions and fill genuinely merged leadership cells only.
    for(let i=0;i<grid.length;i++)grid[i]??=[];
    for(const m of elements(doc,'mergeCell')){
      const [a,b]=m.getAttribute('ref').split(':');if(!b)continue;const r1=+a.match(/\d+/)[0]-1,r2=+b.match(/\d+/)[0]-1,c1=col(a),c2=col(b);
      if(c1===c2&&r2-r1<1000)for(let r=r1+1;r<=r2;r++) {grid[r]??=[];grid[r][c1]=grid[r1]?.[c1]??'';}
    }
    sheets.push({name:sheet.getAttribute('name'),grid});
  }
  return sheets;
}
