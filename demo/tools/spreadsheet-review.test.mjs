import test from 'node:test';
import assert from 'node:assert/strict';
import {sheetAxis, columnName, cellPosition, indexReview, sourceKey, revealIn} from '../browser/wwwroot/spreadsheet-review.js';
import {readFileSync,writeFileSync,mkdirSync} from 'node:fs';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {zipSync,strToU8} from 'fflate';

test('changes beyond the previous 60-row/20-column cutoff remain visible with context', () => {
  const rows = sheetAxis({max:10000,changed:[9000]});
  const columns = sheetAxis({max:100,changed:[27],pageSize:24});
  assert.deepEqual(rows.values,[1,8999,9000,9001]);
  assert.ok(columns.values.includes(27));
  assert.ok(rows.entries.some(entry=>entry.start===2&&entry.end===8998));
  assert.equal(columnName(27),'AA');
  assert.equal(columnName(16384),'XFD');
  assert.deepEqual(cellPosition({locator:{address:'XFD1048576'}}),{row:1048576,column:16384});
});

test('large sets of changes are paginated and direct selection opens the right page', () => {
  const changed=Array.from({length:10000},(_,i)=>i+1);
  const initial=sheetAxis({max:10000,changed});
  const selected=sheetAxis({max:10000,changed,target:9999});
  assert.equal(initial.values.length,80);
  assert.ok(!initial.values.includes(9999));
  assert.ok(selected.values.includes(9999));
  assert.ok(selected.values.length<=80);
  assert.equal(selected.pages,125);
});

test('full-sheet browsing stays bounded even at the Excel row and column limits', () => {
  const rows=sheetAxis({max:1048576,mode:'all',target:1048576});
  const columns=sheetAxis({max:16384,mode:'all',pageSize:24,target:16384});
  assert.ok(rows.values.includes(1048576));
  assert.ok(rows.values.length<=80);
  assert.ok(columns.values.includes(16384));
  assert.ok(columns.values.length<=24);
  assert.equal(rows.hidden,0);
  assert.equal(columns.hidden,0);
});

test('expanding hidden context reveals the requested rows without losing distant changes', () => {
  const expanded=sheetAxis({max:10000,changed:[9000],expanded:[[2,51]]});
  assert.ok(expanded.values.includes(30));
  assert.ok(expanded.values.includes(9000));
  assert.ok(expanded.entries.some(entry=>entry.start===52&&entry.end===8998));
});

test('empty sheets, unchanged sheets, first/last rows and invalidated pages have usable windows', () => {
  assert.deepEqual(sheetAxis({max:1}).values,[1]);
  assert.deepEqual(sheetAxis({max:100}).values,[1,2,3,4,5,6,7,8]);
  assert.deepEqual(sheetAxis({max:10,changed:[1,10],page:99}).values,[1,2,9,10]);
  assert.equal(cellPosition(null),null);
});

test('before/after sources retain independent coordinates after row insertion; additions have no false old anchor', () => {
  const source=address=>({part:'/xl/worksheets/sheet1.xml',locator:{type:'excel',sheetId:1,address}});
  const before={source:source('AA250')},after={source:source('AA251')};
  const change={id:'updated',before,after},added={id:'added',before:null,after:{source:source('AF280')}};
  const result={changeSet:{changes:[change,added]},before:{nodes:[{pageKey:'sheet:1',row:250,column:27,source:before.source}]},after:{nodes:[{pageKey:'sheet:1',row:251,column:27,source:after.source},{pageKey:'sheet:1',row:280,column:32,source:added.after.source}]}};
  const index=indexReview(result);
  assert.equal(index.changes.before.get(sourceKey(source('AA250')))[0].id,'updated');
  assert.equal(index.changes.after.get(sourceKey(source('AA251')))[0].id,'updated');
  assert.equal(index.changes.before.has(sourceKey(source('AF280'))),false);
  assert.equal(index.sheets.after.get('sheet:1').maxColumn,32);
  assert.equal(index.sheets.after.get('sheet:1').maxRow,280);
});

test('reveal centers within the preview without asking the document or another pane to scroll', () => {
  let scroll;
  const container={scrollTop:500,scrollLeft:100,clientHeight:400,clientWidth:300,getBoundingClientRect:()=>({top:100,left:50}),scrollTo:value=>{scroll=value;}};
  revealIn(container,{getBoundingClientRect:()=>({top:800,left:750,height:36,width:150})},'smooth');
  assert.deepEqual(scroll,{top:1018,left:725,behavior:'smooth'});
});

test('real multi-sheet XLSX supports remote cells, pagination, inserted rows, additions and deletions',async()=>{
  globalThis.DOMParser=DOMParser;globalThis.XMLSerializer=XMLSerializer;
  const {OfficeEngine}=await import('../browser/wwwroot/engine/office-engine.js');
  const {CharacterDiffer}=await import('../browser/wwwroot/engine/character-diff.js');
  const differ=await CharacterDiffer.load(readFileSync(new URL('../browser/wwwroot/character-diff.wasm',import.meta.url)));
  const ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main',rel='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const text=(address,value)=>`<c r="${address}" t="inlineStr"><is><t>${value}</t></is></c>`;
  const number=(address,value)=>`<c r="${address}"><v>${value}</v></c>`;
  function workbook(side){
    const next=side==='after';let rows='<row r="1">'+['A','B','C','AA','AF'].map((c,i)=>text(c+'1',['部门','预算','备注','远端预算','状态'][i])).join('')+'</row>';
    for(let r=2;r<=360;r++){
      const address=r+(next&&r>=100?1:0);
      rows+=`<row r="${address}">${text('A'+address,'部门'+r)}${number('B'+address,next&&r%3===0?150:120)}${text('C'+address,'核对上下文')}${number('AA'+address,next&&r===350?1500000:1200000)}${next&&r===300?'':text('AF'+address,'有效')}</row>`;
    }
    if(next)rows+='<row r="100">'+text('A100','新增部门')+number('B100',300)+'</row>';
    const sheet=data=>`<worksheet xmlns="${ns}"><sheetData>${data}</sheetData></worksheet>`;
    const files={
      '[Content_Types].xml':'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>',
      '_rels/.rels':`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
      'xl/workbook.xml':`<workbook xmlns="${ns}" xmlns:r="${rel}"><sheets><sheet name="大表预算" sheetId="1" r:id="rId1"/><sheet name="远端工作表" sheetId="2" r:id="rId2"/></sheets></workbook>`,
      'xl/_rels/workbook.xml.rels':`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${rel}/worksheet" Target="worksheets/sheet2.xml"/></Relationships>`,
      'xl/worksheets/sheet1.xml':sheet(rows),
      'xl/worksheets/sheet2.xml':sheet('<row r="1">'+text('A1','项目')+text('AA1','金额')+'</row><row r="9000">'+text('A9000','远端项目')+number('AA9000',next?150:120)+'</row>')
    };
    return zipSync(Object.fromEntries(Object.entries(files).map(([name,xml])=>[name,strToU8(xml)])));
  }
  const before=workbook('before'),after=workbook('after');
  const result=await new OfficeEngine(differ).compare(before,after,'large-review-before.xlsx','large-review-after.xlsx');
  const index=indexReview(result);
  const moved=result.changeSet.changes.find(change=>change.before?.source.locator.address==='AA350');
  assert.equal(moved.after.source.locator.address,'AA351');
  assert.equal(result.changeSet.changes.filter(c=>c.before?.source.locator.sheetId===2).length,1);
  assert.ok(result.changeSet.changes.some(c=>c.operation==='add'&&c.after.source.locator.address==='A100'));
  assert.ok(result.changeSet.changes.some(c=>c.operation==='remove'&&c.before.source.locator.address==='AF300'));
  const rows=sheetAxis({max:index.sheets.after.get('sheet:1').maxRow,changed:result.changeSet.changes.filter(c=>c.after?.source.locator.sheetId===1).map(c=>cellPosition(c.after.source).row),target:351});
  assert.ok(rows.pages>1);assert.ok(rows.values.includes(351));assert.ok(rows.values.length<=80);
  mkdirSync(new URL('../artifacts/',import.meta.url),{recursive:true});
  for(const [side,bytes]of [['before',before],['after',after]])writeFileSync(new URL(`../artifacts/large-review-${side}.xlsx`,import.meta.url),bytes);
});
