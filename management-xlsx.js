const XML_HEADER='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const MIME='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const CHART_COLORS=['CB2639','087F84','607CF0','EDAA4B','D66AB7','64A5DB','91BF65','EE6570','9383D5','48A881','B28C61','526D82'];
const encoder=new TextEncoder();

const xml=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]));
const finite=v=>v===null||v===undefined||v===''?null:Number.isFinite(Number(v))?Number(v):null;

function columnName(index){
  let name='',value=index+1;
  while(value){value--;name=String.fromCharCode(65+value%26)+name;value=Math.floor(value/26);}
  return name;
}

function textCell(reference,value,style=0){
  const content=String(value??'');
  const preserve=/^\s|\s$/.test(content)?' xml:space="preserve"':'';
  return `<c r="${reference}" s="${style}" t="inlineStr"><is><t${preserve}>${xml(content)}</t></is></c>`;
}

function numberCell(reference,value,style=0){
  const n=finite(value);
  return n===null?`<c r="${reference}" s="${style}"/>`:`<c r="${reference}" s="${style}"><v>${n}</v></c>`;
}

function rowXml(number,cells,height){
  return `<row r="${number}"${height?` ht="${height}" customHeight="1"`:''}>${cells.join('')}</row>`;
}

function dosDateTime(date){
  const d=date instanceof Date&&!Number.isNaN(date.getTime())?date:new Date();
  const year=Math.max(1980,d.getFullYear());
  return {date:((year-1980)<<9)|((d.getMonth()+1)<<5)|d.getDate(),time:(d.getHours()<<11)|(d.getMinutes()<<5)|(d.getSeconds()>>1)};
}

const crcTable=(()=>{
  const table=new Uint32Array(256);
  for(let n=0;n<256;n++){let c=n;for(let k=0;k<8;k++)c=(c&1)?0xEDB88320^(c>>>1):c>>>1;table[n]=c>>>0;}
  return table;
})();

function crc32(bytes){
  let crc=0xFFFFFFFF;
  for(const byte of bytes)crc=crcTable[(crc^byte)&0xFF]^(crc>>>8);
  return (crc^0xFFFFFFFF)>>>0;
}

function writeU16(view,offset,value){view.setUint16(offset,value,true);}
function writeU32(view,offset,value){view.setUint32(offset,value>>>0,true);}

function zip(files,stamp){
  const records=[];
  let localSize=0;
  const when=dosDateTime(stamp);
  for(const [name,content] of files){
    const nameBytes=encoder.encode(name);const data=typeof content==='string'?encoder.encode(content):content;const crc=crc32(data);
    records.push({nameBytes,data,crc,offset:localSize});localSize+=30+nameBytes.length+data.length;
  }
  const centralSize=records.reduce((sum,r)=>sum+46+r.nameBytes.length,0);
  const output=new Uint8Array(localSize+centralSize+22);const view=new DataView(output.buffer);let offset=0;
  for(const r of records){
    writeU32(view,offset,0x04034B50);writeU16(view,offset+4,20);writeU16(view,offset+6,0x0800);writeU16(view,offset+8,0);writeU16(view,offset+10,when.time);writeU16(view,offset+12,when.date);writeU32(view,offset+14,r.crc);writeU32(view,offset+18,r.data.length);writeU32(view,offset+22,r.data.length);writeU16(view,offset+26,r.nameBytes.length);writeU16(view,offset+28,0);output.set(r.nameBytes,offset+30);output.set(r.data,offset+30+r.nameBytes.length);offset+=30+r.nameBytes.length+r.data.length;
  }
  const centralOffset=offset;
  for(const r of records){
    writeU32(view,offset,0x02014B50);writeU16(view,offset+4,20);writeU16(view,offset+6,20);writeU16(view,offset+8,0x0800);writeU16(view,offset+10,0);writeU16(view,offset+12,when.time);writeU16(view,offset+14,when.date);writeU32(view,offset+16,r.crc);writeU32(view,offset+20,r.data.length);writeU32(view,offset+24,r.data.length);writeU16(view,offset+28,r.nameBytes.length);writeU16(view,offset+30,0);writeU16(view,offset+32,0);writeU16(view,offset+34,0);writeU16(view,offset+36,0);writeU32(view,offset+38,0);writeU32(view,offset+42,r.offset);output.set(r.nameBytes,offset+46);offset+=46+r.nameBytes.length;
  }
  writeU32(view,offset,0x06054B50);writeU16(view,offset+4,0);writeU16(view,offset+6,0);writeU16(view,offset+8,records.length);writeU16(view,offset+10,records.length);writeU32(view,offset+12,centralSize);writeU32(view,offset+16,centralOffset);writeU16(view,offset+20,0);
  return output;
}

function contentTypes(hasChart){
  return XML_HEADER+`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${hasChart?'<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>':''}<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`;
}

const packageRels=XML_HEADER+'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>';
const workbookXml=XML_HEADER+'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="27328"/><workbookPr defaultThemeVersion="166925"/><bookViews><workbookView xWindow="0" yWindow="0" windowWidth="24000" windowHeight="13500"/></bookViews><sheets><sheet name="Leader-Wise" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="191029" fullCalcOnLoad="1" forceFullCalc="1"/></workbook>';
const workbookRels=XML_HEADER+'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>';

const stylesXml=XML_HEADER+`<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="2"><numFmt numFmtId="164" formatCode="#,##0.00;[Red]-#,##0.00;–"/><numFmt numFmtId="165" formatCode="0.0%"/></numFmts><fonts count="5"><font><sz val="10"/><color rgb="FF142235"/><name val="Arial"/><family val="2"/></font><font><b/><sz val="18"/><color rgb="FF142235"/><name val="Arial"/><family val="2"/></font><font><i/><sz val="10"/><color rgb="FF586C83"/><name val="Arial"/><family val="2"/></font><font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Arial"/><family val="2"/></font><font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Arial"/><family val="2"/></font></fonts><fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF142235"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFCB2639"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FFDCE3ED"/></left><right style="thin"><color rgb="FFDCE3ED"/></right><top style="thin"><color rgb="FFDCE3ED"/></top><bottom style="thin"><color rgb="FFDCE3ED"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="12"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="0" fontId="3" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf><xf numFmtId="3" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf><xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf><xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf><xf numFmtId="0" fontId="4" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf><xf numFmtId="3" fontId="4" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf><xf numFmtId="164" fontId="4" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf><xf numFmtId="165" fontId="4" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/></styleSheet>`;

function worksheetXml(model,hasChart){
  const headers=['RHO / Leader','Outlets',...model.banks.map(b=>b.label),'Total Extra Amount','Month End Projection','Share'];
  const lastColumn=columnName(headers.length-1);const firstDataRow=6;const totalRow=firstDataRow+model.rows.length;
  const sheetRows=[
    rowXml(1,[textCell('A1','Credit Card Extra Amount — Leader-Wise Management',1)],28),
    rowXml(2,[textCell('A2',`Reporting period: ${model.period||'Not provided'}`,2)],18),
    rowXml(3,[textCell('A3',`Scope: ${model.scope||'All active outlets'}`,2)],18),
    rowXml(4,[textCell('A4',`Generated: ${model.generatedLabel||''}`,2)],18),
    rowXml(5,headers.map((header,index)=>textCell(`${columnName(index)}5`,header,3)),34)
  ];
  model.rows.forEach((row,index)=>{
    const n=firstDataRow+index;const values=[textCell(`A${n}`,row.leader,4),numberCell(`B${n}`,row.outlets,5)];
    row.channels.forEach((value,channelIndex)=>values.push(numberCell(`${columnName(2+channelIndex)}${n}`,value,6)));
    const extraColumn=2+model.banks.length;values.push(numberCell(`${columnName(extraColumn)}${n}`,row.extra,6),numberCell(`${columnName(extraColumn+1)}${n}`,row.projection,6),numberCell(`${columnName(extraColumn+2)}${n}`,row.share,7));
    sheetRows.push(rowXml(n,values,19));
  });
  const totalCells=[textCell(`A${totalRow}`,'VIEW TOTAL',8),numberCell(`B${totalRow}`,model.totals.outlets,9)];
  model.totals.channels.forEach((value,index)=>totalCells.push(numberCell(`${columnName(2+index)}${totalRow}`,value,10)));
  const extraColumn=2+model.banks.length;totalCells.push(numberCell(`${columnName(extraColumn)}${totalRow}`,model.totals.extra,10),numberCell(`${columnName(extraColumn+1)}${totalRow}`,model.totals.projection,10),numberCell(`${columnName(extraColumn+2)}${totalRow}`,model.totals.share,11));sheetRows.push(rowXml(totalRow,totalCells,21));
  const widths=headers.map((header,index)=>{const width=index===0?27:index===1?11:index>=headers.length-3?(index===headers.length-1?12:22):Math.min(19,Math.max(12,String(header).length+3));return `<col min="${index+1}" max="${index+1}" width="${width}" customWidth="1"/>`;}).join('');
  const merges=`<mergeCells count="4"><mergeCell ref="A1:${lastColumn}1"/><mergeCell ref="A2:${lastColumn}2"/><mergeCell ref="A3:${lastColumn}3"/><mergeCell ref="A4:${lastColumn}4"/></mergeCells>`;
  return XML_HEADER+`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><dimension ref="A1:${lastColumn}${totalRow}"/><sheetViews><sheetView showGridLines="0" tabSelected="1" workbookViewId="0"><pane ySplit="5" topLeftCell="A6" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A6" sqref="A6"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/><cols>${widths}</cols><sheetData>${sheetRows.join('')}</sheetData><autoFilter ref="A5:${lastColumn}${Math.max(5,totalRow-1)}"/>${merges}<pageMargins left="0.25" right="0.25" top="0.5" bottom="0.5" header="0.2" footer="0.2"/><pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0" paperSize="9"/>${hasChart?'<drawing r:id="rId1"/>':''}</worksheet>`;
}

function chartXml(model){
  const firstRow=6,lastRow=5+model.rows.length,totalExtraColumn=columnName(2+model.banks.length);const sheet="'Leader-Wise'";
  const categoryCache=model.rows.map((row,index)=>`<c:pt idx="${index}"><c:v>${xml(row.leader)}</c:v></c:pt>`).join('');
  const valueCache=model.rows.map((row,index)=>`<c:pt idx="${index}"><c:v>${finite(row.extra)??0}</c:v></c:pt>`).join('');
  const points=model.rows.map((_,index)=>`<c:dPt><c:idx val="${index}"/><c:spPr><a:solidFill><a:srgbClr val="${CHART_COLORS[index%CHART_COLORS.length]}"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr></c:dPt>`).join('');
  return XML_HEADER+`<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="1400" b="1"><a:solidFill><a:srgbClr val="142235"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>Total Extra Amount by RHO / Leader</a:t></a:r></a:p></c:rich></c:tx><c:layout/><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:pieChart><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Total Extra Amount</c:v></c:tx>${points}<c:cat><c:strRef><c:f>${sheet}!$A$${firstRow}:$A$${lastRow}</c:f><c:strCache><c:ptCount val="${model.rows.length}"/>${categoryCache}</c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>${sheet}!$${totalExtraColumn}$${firstRow}:$${totalExtraColumn}$${lastRow}</c:f><c:numCache><c:formatCode>#,##0.00</c:formatCode><c:ptCount val="${model.rows.length}"/>${valueCache}</c:numCache></c:numRef></c:val></c:ser><c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="1"/><c:showSerName val="0"/><c:showPercent val="1"/><c:showLeaderLines val="1"/><c:separator>\n</c:separator></c:dLbls><c:firstSliceAng val="270"/></c:pieChart></c:plotArea><c:legend><c:legendPos val="r"/><c:layout/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/><c:showDLblsOverMax val="0"/></c:chart><c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings></c:chartSpace>`;
}

function drawingXml(columnCount){
  const startColumn=columnCount+1,endColumn=startColumn+9;
  return XML_HEADER+`<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>${startColumn}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>${endColumn}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>21</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Leader-Wise Pie Chart"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`;
}

const drawingRels=XML_HEADER+'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>';
const sheetRels=XML_HEADER+'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>';

function appProperties(){return XML_HEADER+'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel Compatible</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>Leader-Wise</vt:lpstr></vt:vector></TitlesOfParts><Company>SHWAPNO</Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion></Properties>';}
function coreProperties(stamp){const value=(stamp instanceof Date&&!Number.isNaN(stamp.getTime())?stamp:new Date()).toISOString();return XML_HEADER+`<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Credit Card Extra Amount — Leader-Wise Management</dc:title><dc:subject>Leader-wise payment channel management export</dc:subject><dc:creator>SHWAPNO Operations Intelligence</dc:creator><cp:lastModifiedBy>SHWAPNO Operations Intelligence</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">${value}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${value}</dcterms:modified></cp:coreProperties>`;}

export function createManagementWorkbook(model){
  const stamp=model.generatedAt instanceof Date?model.generatedAt:new Date();
  const normalized={...model,banks:Array.isArray(model.banks)?model.banks:[],rows:Array.isArray(model.rows)?model.rows:[],totals:model.totals||{outlets:0,channels:[],extra:0,projection:null,share:0}};
  const hasChart=normalized.rows.length>0;const files=[
    ['[Content_Types].xml',contentTypes(hasChart)],['_rels/.rels',packageRels],['docProps/app.xml',appProperties()],['docProps/core.xml',coreProperties(stamp)],['xl/workbook.xml',workbookXml],['xl/_rels/workbook.xml.rels',workbookRels],['xl/styles.xml',stylesXml],['xl/worksheets/sheet1.xml',worksheetXml(normalized,hasChart)]
  ];
  if(hasChart)files.push(['xl/worksheets/_rels/sheet1.xml.rels',sheetRels],['xl/drawings/drawing1.xml',drawingXml(5+normalized.banks.length)],['xl/drawings/_rels/drawing1.xml.rels',drawingRels],['xl/charts/chart1.xml',chartXml(normalized)]);
  return new Blob([zip(files,stamp)],{type:MIME});
}
