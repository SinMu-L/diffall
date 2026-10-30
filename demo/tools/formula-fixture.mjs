import {zipSync,strToU8} from 'fflate';
const ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main',rel='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const escape=value=>String(value).replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;');
const text=(address,value)=>`<c r="${address}" t="inlineStr"><is><t>${escape(value)}</t></is></c>`;
const number=(address,value)=>`<c r="${address}"><v>${value}</v></c>`;
const formula=(address,expression,value)=>`<c r="${address}"><f>${escape(expression)}</f>${value==null?'':`<v>${value}</v>`}</c>`;
const row=(index,cells)=>`<row r="${index}">${cells}</row>`;
const sheet=rows=>`<worksheet xmlns="${ns}"><sheetData>${rows.join('')}</sheetData></worksheet>`;
export function createFormulaWorkbook(next=false){
 const rows=[row(1,text('A1','项目')+text('B1','金额')+text('D1','计算结果')),
 row(2,text('A2','范围合计')+formula('D2','SUM(B5:B100)',next?110:100)),
 row(3,text('A3','合计倍增')+formula('D3','D2*2',next?220:200)),
 row(4,text('A4','结果缓存相同')+formula('D4','IF(B5>0,0,1)',0)),
 row(5,text('A5','变动输入')+number('B5',next?20:10)+formula('D5','CUSTOM(B5)',null)),
 row(6,text('A6','固定输入与公式修改')+number('B6',90)+formula('D6',next?'SUM(B5:B100)+1':'SUM(B5:B100)',next?111:100)),row(100,text('A100','范围末行')+number('B100',0))];
 const files={
 '[Content_Types].xml':`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
 '_rels/.rels':`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
 'xl/workbook.xml':`<workbook xmlns="${ns}" xmlns:r="${rel}"><sheets><sheet name="预算明细" sheetId="1" r:id="rId1"/><sheet name="汇总" sheetId="2" r:id="rId2"/></sheets><calcPr calcMode="auto" fullCalcOnLoad="1"/></workbook>`,
 'xl/_rels/workbook.xml.rels':`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${rel}/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="${rel}/styles" Target="styles.xml"/></Relationships>`,
 'xl/styles.xml':`<styleSheet xmlns="${ns}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>`,
 'xl/worksheets/sheet1.xml':sheet(rows),
 'xl/worksheets/sheet2.xml':sheet([row(1,text('A1','项目')+text('B1','汇总金额')),row(2,text('A2','跨表结果')+formula('B2',"'预算明细'!D3",next?220:200))])};
 return zipSync(Object.fromEntries(Object.entries(files).map(([name,xml])=>[name,strToU8(xml)])));
}
