import test from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {createFormulaWorkbook} from './formula-fixture.mjs';
import {characterDiff} from '../browser/wwwroot/text-diff-view.js';
import {formatNumberText} from '../browser/wwwroot/value-format.js';

globalThis.DOMParser=DOMParser;globalThis.XMLSerializer=XMLSerializer;
const {OfficeEngine}=await import('../browser/wwwroot/engine/office-engine.js');
const {CharacterDiffer}=await import('../browser/wwwroot/engine/character-diff.js');
const {OfficePackage,descendants,first,children,attr,append,remove}=await import('../browser/wwwroot/engine/ooxml.js');
const differ=await CharacterDiffer.load(readFileSync(new URL('../browser/wwwroot/character-diff.wasm',import.meta.url))),engine=new OfficeEngine(differ);
const part='xl/worksheets/sheet1.xml';
const inputCell=pkg=>descendants(pkg.xml(part),'s','c').find(cell=>attr(cell,'r')==='B5');
function workbook(data,type='number'){
 const pkg=new OfficePackage(createFormulaWorkbook()),cell=inputCell(pkg);
 for(const child of children(cell))remove(child);
 if(type==='text'){cell.setAttribute('t','inlineStr');append(append(cell,'s','is'),'s','t',{},data);}
 else{cell.removeAttribute('t');append(cell,'s','v',{},data);}
 pkg.mark(part);return pkg.write();
}

test('numeric values produce whole-value changes and native cell highlights without invoking character alignment',async t=>{
 for(const [old,next] of [['120','150'],['999','1000'],['10','-10'],['0.01','0.001'],['300.1234567890123456','300.1234567890123457'],['9007199254740992','9007199254740993'],['1.2e20','1.5e20']]){
  await t.test(`${old} -> ${next}`,async()=>{
   const before=workbook(old),after=workbook(next),calls=differ.calls;
   const result=await engine.compare(before,after,'old.xlsx','new.xlsx');
   assert.equal(differ.calls,calls,'numeric input changes do not run the character differ');
   assert.equal(result.changeSet.changes.length,1);
   const change=result.changeSet.changes[0];
   assert.equal(change.aspect,'value');assert.equal(change.extensions?.['ppt2wiki.textDiff'],undefined);assert.equal(characterDiff(change),undefined);
   assert.deepEqual(change.before.value,{type:'number',data:old});assert.deepEqual(change.after.value,{type:'number',data:next});
   for(const [side,bytes,value] of [['before',before,old],['after',after,next]]){
    const cell=inputCell(new OfficePackage(await engine.annotate(bytes,result.changeSet,side)));
    assert.equal(first(cell,'s','v').textContent,value);assert.equal(first(cell,'s','is'),undefined);
    assert.notEqual(attr(cell,'t'),'inlineStr');assert.ok(Number(attr(cell,'s'))>0,'the complete numeric cell is highlighted');
   }
  });
 }
});

test('equivalent decimal representations do not produce numeric changes',async()=>{
 for(const [old,next] of [['1','1.00'],['1200000','1.200e6'],['-0','0'],['0.00100','1e-3']]){
  const result=await engine.compare(workbook(old),workbook(next),'old.xlsx','new.xlsx');
  assert.deepEqual(result.changeSet.changes,[]);
 }
});

test('numeric identifiers remain text with leading zeroes and character-level edits',async()=>{
 const result=await engine.compare(workbook('00120','text'),workbook('00150','text'),'old.xlsx','new.xlsx');
 assert.equal(result.changeSet.changes.length,1);
 const change=result.changeSet.changes[0],diff=characterDiff(change);
 assert.deepEqual(change.before.value,{type:'text',data:'00120'});assert.deepEqual(change.after.value,{type:'text',data:'00150'});
 assert.equal(diff.edits.length,1);assert.equal(diff.edits[0].before.text,'2');assert.equal(diff.edits[0].after.text,'5');
});

test('number/text type changes remain whole-value updates even when the characters are identical',async()=>{
 for(const [before,after] of [[workbook('120'),workbook('120','text')],[workbook('120','text'),workbook('120')]]){
  const result=await engine.compare(before,after,'old.xlsx','new.xlsx');
  assert.equal(result.changeSet.changes.length,1);
  const change=result.changeSet.changes[0];
  assert.notEqual(change.before.value.type,change.after.value.type);assert.equal(change.extensions?.['ppt2wiki.textDiff'],undefined);
 }
});

test('formula results use whole numeric values while formula expressions retain character edits',async()=>{
 const result=await engine.compare(createFormulaWorkbook(),createFormulaWorkbook(true),'old.xlsx','new.xlsx');
 const values=result.changeSet.changes.filter(change=>change.aspect==='value'),formulas=result.changeSet.changes.filter(change=>change.aspect==='formula');
 assert.ok(values.length>0);assert.ok(formulas.length>0);
 for(const change of values){assert.equal(change.before.value.type,'number');assert.equal(change.extensions?.['ppt2wiki.textDiff'],undefined);}
 for(const change of formulas)assert.ok(characterDiff(change).edits.length>0);
});

test('preview ignores legacy numeric character ranges and retains text/formula ranges',()=>{
 const diff=differ.compare('120','150'),change={aspect:'value',before:{value:{type:'number',data:'120'}},after:{value:{type:'number',data:'150'}},extensions:{'ppt2wiki.textDiff':diff}};
 assert.equal(characterDiff(change),undefined);
 change.before.value.type='text';assert.equal(characterDiff(change),undefined);
 change.after.value.type='text';assert.equal(characterDiff(change),diff);
 change.aspect='formula';assert.equal(characterDiff(change),diff);
});

test('number presentation keeps every significant digit, decimal place and exponent',()=>{
 for(const [raw,display] of [['1200000','1,200,000'],['-1234567.1234567890123456','-1,234,567.1234567890123456'],['300.1234567890123456','300.1234567890123456'],['300.1234567890123457','300.1234567890123457'],['9007199254740992','9,007,199,254,740,992'],['9007199254740993','9,007,199,254,740,993'],['0.0000000001','0.0000000001'],['1.00','1.00'],['-0','-0'],['1.234567890123456789e+20','1.234567890123456789e+20']]){
  assert.equal(formatNumberText(raw),display);assert.equal(display.replaceAll(',',''),raw);
 }
});
