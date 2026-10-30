import test from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync,writeFileSync,mkdirSync} from 'node:fs';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {createFormulaWorkbook} from './formula-fixture.mjs';
import {parseFormulaReferences} from '../browser/wwwroot/engine/formula-impact.js';
import {indexImpacts,impactLabel,relatedPositions,referencedAreas} from '../browser/wwwroot/impact-review.js';
import {sheetAxis} from '../browser/wwwroot/spreadsheet-review.js';
globalThis.DOMParser=DOMParser;globalThis.XMLSerializer=XMLSerializer;
const {OfficeEngine}=await import('../browser/wwwroot/engine/office-engine.js');
const {CharacterDiffer}=await import('../browser/wwwroot/engine/character-diff.js');
const {OfficePackage,descendants,first,attr}=await import('../browser/wwwroot/engine/ooxml.js');
const differ=await CharacterDiffer.load(readFileSync(new URL('../browser/wwwroot/character-diff.wasm',import.meta.url))),engine=new OfficeEngine(differ);
const compare=()=>engine.compare(createFormulaWorkbook(),createFormulaWorkbook(true),'formulas-old.xlsx','formulas-new.xlsx');
const target=(result,address,sheetId=1)=>result.after.nodes.find(n=>n.source.locator.sheetId===sheetId&&n.source.locator.address===address);

test('static references preserve absolute/range/cross-sheet addresses and exclude strings and scientific constants',()=>{
 const parsed=parseFormulaReferences(`SUM($A$1:B10)+'预算 表'!$C$2+1E10+"D9"+LOG10(A2)`);
 assert.deepEqual(parsed.references.map(r=>[r.sheet,r.start,r.end]),[[null,'A1','B10'],['预算 表','C2','C2'],[null,'A2','A2']]);
 assert.ok(!parsed.references.some(r=>r.start==='E10'||r.start==='D9'||r.start==='LOG10'));
 assert.equal(parseFormulaReferences('SUM(A:A)+SUM(2:5)').references.length,2);
});
test('dynamic, named and external references do not create a false local input',()=>{
 assert.ok(parseFormulaReferences('INDIRECT("A1")').issues.length);
 assert.ok(parseFormulaReferences('SUM(BudgetRange)').issues.length);
 const external=parseFormulaReferences('[Other.xlsx]汇总!B2+\'[Book.xlsx]汇总\'!A1');
 assert.equal(external.references.length,0);assert.ok(external.issues.length);
});
test('saved formula results produce real value updates with unchanged formulas, range and cross-sheet causal candidates',async()=>{
 const result=await compare(),index=indexImpacts(result),b=target(result,'D2'),root=target(result,'B5'),cross=target(result,'B2',2);
 assert.equal(result.changeSet.changes.length,6);assert.equal(index.stats.changed,4);assert.equal(index.stats.unknown,2);
 const change=result.changeSet.changes.find(c=>c.target.entityId===b.entityId&&c.aspect==='value');
 assert.equal(change.before.value.data,'100');assert.equal(change.after.value.data,'110');assert.equal(index.byChange.get(change.id).evidence.level,'supported');
 assert.ok(!result.changeSet.changes.some(c=>c.target.entityId===b.entityId&&c.aspect==='formula'));
 assert.equal(index.relatedDownstream(root.entityId).length,6);
 assert.ok(index.byEntity.get(cross.entityId).causes.some(c=>c.dependencyIds.length));
 assert.ok(referencedAreas(index,b.entityId,'after').some(a=>a.start==='B5'&&a.end==='B100'));
 assert.equal(result.changeSet.comparison.compareCachedFormulaValues,true);
 mkdirSync(new URL('../artifacts/',import.meta.url),{recursive:true});
 writeFileSync(new URL('../artifacts/formula-impact.changes.json',import.meta.url),JSON.stringify(result.changeSet,null,2));
});
test('equal and unavailable cache results remain unknown, never generate fake updates or inflate actual changes',async()=>{
 const result=await compare(),index=indexImpacts(result);
 for(const address of ['D4','D5']){const node=target(result,address),impact=index.byEntity.get(node.entityId);assert.equal(impact.outcome,'unknown');assert.deepEqual(impact.changeIds,[]);assert.equal(impactLabel(impact),'结果待确认');assert.ok(!result.changeSet.changes.some(c=>c.target.entityId===node.entityId));}
 assert.equal(target(result,'D5').cachedValue,null);
 assert.equal(index.items.length,6);assert.equal(result.changeSet.changes.length,6);
});
test('unrelated input edits leave unresolved formulas in analysis limitations rather than calculation impacts',async t=>{
 for(const [name,expression,shared] of [['unsupported function','CUSTOM(B5)'],['named range','SUM(BudgetRange)'],['dynamic reference','INDIRECT("B5")'],['external reference',"'[Other.xlsx]汇总'!B2"],['shared formula','',true]]){
  await t.test(name,async()=>{
   function workbook(input){const pkg=new OfficePackage(createFormulaWorkbook()),root=pkg.xml('xl/worksheets/sheet1.xml'),cells=descendants(root,'s','c'),f=first(cells.find(c=>attr(c,'r')==='D4'),'s','f');f.textContent=expression;if(shared){f.setAttribute('t','shared');f.setAttribute('si','0');}first(cells.find(c=>attr(c,'r')==='B6'),'s','v').textContent=input;pkg.mark('xl/worksheets/sheet1.xml');return pkg.write();}
   const result=await engine.compare(workbook('90'),workbook('95'),'a.xlsx','b.xlsx'),index=indexImpacts(result),unrelated=target(result,'D4'),missing=target(result,'D5');
   assert.equal(result.changeSet.changes.length,1);assert.equal(result.changeSet.changes[0].after.source.locator.address,'B6');
   assert.equal(index.byEntity.has(unrelated.entityId),false);assert.equal(index.byEntity.has(missing.entityId),false);
   assert.deepEqual(index.items.map(i=>i.after.source.locator.address).sort(),['B2','D2','D3','D6']);
   assert.equal(index.stats.changed,0);assert.equal(index.stats.unknown,4);
   assert.equal(index.graph.coverage.dependencies.status,'partial');
   assert.ok(index.graph.coverage.dependencies.issues.some(i=>i.entityIds.includes(unrelated.entityId)));
   assert.equal(index.dependencyLimitations.count,2);assert.equal(new Set(index.dependencyLimitations.details).size,index.dependencyLimitations.details.length);
  });
 }
});

test('style-only edits do not turn unresolved formulas into calculation impacts',async()=>{
 const before=createFormulaWorkbook(),pkg=new OfficePackage(before),root=pkg.xml('xl/styles.xml'),font=descendants(root,'s','font')[0];font.appendChild(font.ownerDocument.createElementNS(font.namespaceURI,'b'));pkg.mark('xl/styles.xml');
 const result=await engine.compare(before,pkg.write(),'a.xlsx','b.xlsx'),index=indexImpacts(result);
 assert.ok(result.changeSet.changes.length>0);assert.ok(result.changeSet.changes.every(c=>c.aspect==='style'));
 assert.equal(index.items.length,0);assert.equal(index.stats.unknown,0);assert.equal(index.dependencyLimitations.count,1);
 assert.equal(index.graph.coverage.dependencies.status,'partial');
});

test('formula edits with equal saved values remain calculation impacts without inventing value changes',async()=>{
 const before=createFormulaWorkbook(),pkg=new OfficePackage(before),root=pkg.xml('xl/worksheets/sheet1.xml'),cell=descendants(root,'s','c').find(c=>attr(c,'r')==='D4');first(cell,'s','f').textContent='IF(B5>0,0,2)';pkg.mark('xl/worksheets/sheet1.xml');
 const result=await engine.compare(before,pkg.write(),'a.xlsx','b.xlsx'),index=indexImpacts(result),node=target(result,'D4'),impact=index.byEntity.get(node.entityId);
 assert.equal(result.changeSet.changes.length,1);assert.equal(result.changeSet.changes[0].aspect,'formula');assert.equal(index.items.length,1);
 assert.equal(impact.outcome,'unknown');assert.deepEqual(impact.changeIds,[]);assert.ok(impact.causes.some(c=>c.changeId===result.changeSet.changes[0].id));
 assert.equal(index.dependencyLimitations.count,1);
});

test('unresolved formulas with changed saved results remain calculation impacts even without an input path',async()=>{
 function workbook(value){const pkg=new OfficePackage(createFormulaWorkbook()),root=pkg.xml('xl/worksheets/sheet1.xml'),cell=descendants(root,'s','c').find(c=>attr(c,'r')==='D4');first(cell,'s','f').textContent='INDIRECT("B5")';first(cell,'s','v').textContent=value;pkg.mark('xl/worksheets/sheet1.xml');return pkg.write();}
 const result=await engine.compare(workbook('10'),workbook('20'),'a.xlsx','b.xlsx'),index=indexImpacts(result),node=target(result,'D4'),impact=index.byEntity.get(node.entityId);
 assert.equal(result.changeSet.changes.length,1);assert.equal(index.items.length,1);assert.equal(impact.outcome,'changed');assert.equal(index.stats.changed,1);
 assert.deepEqual(impact.changeIds,[result.changeSet.changes[0].id]);assert.equal(impact.causes[0].kind,'unresolved');assert.equal(index.dependencyLimitations.count,2);
});

test('focused calculation navigation includes unchanged context and inputs even when they are outside the current filter',async()=>{
 const result=await compare(),index=indexImpacts(result),b=target(result,'D2');
 const positions=relatedPositions(index,b.entityId,'after','sheet:1');
 assert.ok(positions.some(p=>p.row===5&&p.column===2));
 const rows=sheetAxis({max:100,changed:positions.map(p=>p.row)});assert.ok(rows.values.includes(5));assert.ok(rows.values.includes(2));
 const cross=target(result,'B2',2),sourceAreas=referencedAreas(index,cross.entityId,'after');assert.equal(sourceAreas.length,1);assert.equal(sourceAreas[0].start,'D3');
 const root=target(result,'B5');assert.equal(index.targetId(root.entityId),result.changeSet.changes.find(c=>c.target.entityId===root.entityId).id);
});
test('formula and value events remain separate facts while calculation counts deduplicate the cell',async()=>{
 const result=await compare(),index=indexImpacts(result),node=target(result,'D6');
 assert.deepEqual(result.changeSet.changes.filter(c=>c.target.entityId===node.entityId).map(c=>c.aspect),['value','formula']);
 const impact=index.byEntity.get(node.entityId);assert.ok(impact.causes.some(c=>c.dependencyIds.length===0));assert.ok(impact.causes.some(c=>c.dependencyIds.length>0));
 assert.equal(index.stats.changed,4);
});
test('formula-result highlighter preserves formulas, exact cache contents and the original snapshot',async()=>{
 const before=createFormulaWorkbook(),after=createFormulaWorkbook(true),result=await engine.compare(before,after,'a.xlsx','b.xlsx');
 for(const [side,bytes] of [['before',before],['after',after]]){const output=await engine.annotate(bytes,result.changeSet,side),original=await engine.extract(bytes,'a.xlsx'),annotated=await engine.extract(output,'a.xlsx');assert.deepEqual(annotated.nodes.map(n=>[n.source,n.formula,n.cachedValue,n.display]),original.nodes.map(n=>[n.source,n.formula,n.cachedValue,n.display]));const pkg=new OfficePackage(output),cell=descendants(pkg.xml('xl/worksheets/sheet1.xml'),'s','c').find(c=>attr(c,'r')==='D2');assert.ok(Number(attr(cell,'s'))>0);}
});
test('numeric first-row inputs match by position and yield value updates instead of false add/remove',async()=>{
 function numeric(next){const pkg=new OfficePackage(createFormulaWorkbook(next)),root=pkg.xml('xl/worksheets/sheet2.xml'),cell=descendants(root,'s','c').find(c=>attr(c,'r')==='A1');cell.removeAttribute('t');while(cell.firstChild)cell.removeChild(cell.firstChild);const v=cell.ownerDocument.createElementNS(root.namespaceURI,'v');v.textContent=next?'20':'10';cell.appendChild(v);pkg.mark('xl/worksheets/sheet2.xml');return pkg.write();}
 const result=await engine.compare(numeric(false),numeric(true),'a.xlsx','b.xlsx');const events=result.changeSet.changes.filter(c=>(c.after||c.before).source.locator.sheetId===2&&(c.after||c.before).source.locator.address==='A1');assert.equal(events.length,1);assert.equal(events[0].aspect,'value');
});
test('an unexpanded shared formula reports missing context and never invents an empty formula update',async()=>{
 function shared(next){const pkg=new OfficePackage(createFormulaWorkbook(next)),root=pkg.xml('xl/worksheets/sheet1.xml'),cell=descendants(root,'s','c').find(c=>attr(c,'r')==='D2'),f=first(cell,'s','f');f.setAttribute('t','shared');f.setAttribute('si','0');f.textContent='';pkg.mark('xl/worksheets/sheet1.xml');return pkg.write();}
 const result=await engine.compare(shared(false),shared(true),'a.xlsx','b.xlsx'),index=indexImpacts(result),b=target(result,'D2');assert.equal(index.byEntity.get(b.entityId).outcome,'changed');assert.ok(!result.changeSet.changes.some(c=>c.target.entityId===b.entityId&&c.aspect==='formula'));assert.equal(index.graph.coverage.dependencies.status,'partial');
 writeFileSync(new URL('../artifacts/formula-shared.changes.json',import.meta.url),JSON.stringify(result.changeSet,null,2));
});
test('unknown extension versions are not interpreted as supported impact facts',async()=>{
 const result=await compare();result.changeSet.extensions['ppt2wiki.impact'].version='99.0.0';const index=indexImpacts(result);assert.equal(index.unsupported,true);assert.equal(index.items.length,0);assert.equal(index.stats.changed,0);
});
