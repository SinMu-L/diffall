import test from 'node:test';
import assert from 'node:assert/strict';
import {readFile,writeFile,mkdir} from 'node:fs/promises';
import {PDFDocument,StandardFonts,rgb,degrees} from 'pdf-lib';
import {createCanvas} from '@napi-rs/canvas';
import {OfficeEngine} from '../browser/wwwroot/engine/office-engine.js';
import {CharacterDiffer} from '../browser/wwwroot/engine/character-diff.js';
import {comparePdfModels,matchObjects,groupSpans} from '../browser/wwwroot/engine/pdf-model.js';
import {createPdfFixture,createCjkPdfFixture} from './pdf-fixture.mjs';
const differ=await CharacterDiffer.load(await readFile(new URL('../browser/wwwroot/character-diff.wasm',import.meta.url)));
const engine=new OfficeEngine(differ),dir=new URL('../artifacts/',import.meta.url);
await mkdir(dir,{recursive:true});
const oldBytes=await createPdfFixture(false),newBytes=await createPdfFixture(true);
const compare=(a,b)=>engine.compare(a,b,'old.pdf','new.pdf');

test('PDF real documents: precise text, mixed styles, position, page insertion/removal and reorder',async()=>{
 const result=await compare(oldBytes,newBytes);
 try{
  const c=result.changeSet.changes;
  assert.equal(result.changeSet.comparison.profile,'pdf-text-v1');
  assert.equal(c.length,9);
  const text=c.find(c=>c.aspect==='text');assert.equal(text.before.value.data,'Approved budget: 100 million');assert.equal(text.after.value.data,'Approved budget: 120 million');
  assert.deepEqual(text.extensions['ppt2wiki.textDiff'].edits.map(e=>[e.before.text,e.after.text]),[['0','2']]);
  assert.deepEqual(c.filter(c=>c.aspect==='style').map(c=>c.property).sort(),['/style/bold','/style/color','/style/fontSize']);
  const bold=c.find(c=>c.property==='/style/bold');assert.ok(bold.after.value.data.some(r=>r.value===true));assert.ok(bold.after.value.data.some(r=>r.value===false));
  assert.equal(c.filter(c=>c.aspect==='geometry').length,1);
  assert.deepEqual(c.filter(c=>c.target.kind==='page').map(c=>[c.operation,c.before?.source.locator.pageIndex,c.after?.source.locator.pageIndex]),[['move',0,2],['remove',2,undefined],['add',undefined,0]]);
  assert.equal(result.before.pages[0].key,result.after.pages[2].key);assert.equal(result.before.pages[1].key,result.after.pages[1].key);
  assert.equal(c.filter(c=>c.operation==='remove').length,1,'whole page deletion must not repeat all descendant removals');
  assert.equal(result.changeSet.coverage.issues.some(i=>i.code==='unresolvedStyle'),false);
  await writeFile(new URL('pdf-budget.changes.json',dir),JSON.stringify(result.changeSet,null,2));
  for(const [side,bytes]of [['before',oldBytes],['after',newBytes]]){
   const out=await engine.annotate(bytes,result.changeSet,side);await writeFile(new URL(`pdf-budget.${side}.highlighted.pdf`,dir),out);
   const original=await engine.extract(bytes,'source.pdf'),annotated=await engine.extract(out,'annotated.pdf');
   try{assert.equal(annotated.pages.length,original.pages.length);assert.deepEqual(annotated.pages.map(p=>p.text),original.pages.map(p=>p.text));}
   finally{await engine.pdf.dispose(original);await engine.pdf.dispose(annotated);}
  }
 }finally{await engine.disposeResult(result);}
});

test('PDF identical snapshots have zero events and input arrays are not detached',async()=>{
 const result=await compare(oldBytes,oldBytes);try{assert.equal(result.changeSet.changes.length,0);assert.ok(oldBytes.byteLength>0);}finally{await engine.disposeResult(result);}
});

async function rearrange(bytes,indices,insert=false){const source=await PDFDocument.load(bytes),doc=await PDFDocument.create();if(insert){const font=await doc.embedFont(StandardFonts.Helvetica);doc.addPage([595,842]).drawText('New independent checklist: 42 tasks',{x:50,y:700,size:15,font});}for(const p of await doc.copyPages(source,indices))doc.addPage(p);return doc.save();}
test('PDF passive page shifts are not moves; actual reorder is one move',async()=>{
 for(const [bytes,expected]of [[await rearrange(oldBytes,[0,1,2],true),[['add','existence']]],[await rearrange(oldBytes,[2,0,1]),[['move','order']]],[await rearrange(oldBytes,[0,2]),[['remove','existence']]]]){
  const result=await compare(oldBytes,bytes);try{assert.deepEqual(result.changeSet.changes.map(c=>[c.operation,c.aspect]),expected);}finally{await engine.disposeResult(result);}
 }
});

test('PDF stale snapshots, edited text anchors and tampered boxes fail before export',async()=>{
 const result=await compare(oldBytes,newBytes);
 try{
  await assert.rejects(engine.annotate(newBytes,result.changeSet,'before'),/快照/);
  for(const mutation of [set=>set.changes.find(c=>c.aspect==='text').after.source.locator.text+='!',set=>set.changes.find(c=>c.aspect==='text').after.source.locator.rect[0]+=10,set=>set.changes.find(c=>c.aspect==='text').after.source.locator.itemIndices=[999999]]){
   const set=structuredClone(result.changeSet);mutation(set);await assert.rejects(engine.annotate(newBytes,set,'after'),/锚点|原始区域/);
  }
 }finally{await engine.disposeResult(result);}
});

test('PDF rotated cropped pages render and exported highlights preserve content and page boxes',async()=>{
 for(const rotation of [90,180,270]){
  const a=await createPdfFixture(false,{rotation,crop:true}),b=await createPdfFixture(true,{rotation,crop:true}),result=await compare(a,b);
  try{
   const page=result.after.pages.find(p=>p.rotation===rotation);assert.ok(page);
   const canvas=createCanvas(1,1);await engine.pdf.render(result.after,page.index,canvas,.6);assert.ok(canvas.width>100&&canvas.height>100);
   const out=await engine.annotate(b,result.changeSet,'after'),doc=await PDFDocument.load(out),original=await PDFDocument.load(b);
   assert.deepEqual(doc.getPages().map(p=>p.getCropBox()),original.getPages().map(p=>p.getCropBox()));assert.equal(doc.getPage(page.index).getRotation().angle,rotation);
  }finally{await engine.disposeResult(result);}
 }
});

test('PDF missing text, invalid bytes and unreadable new revision do not leak sessions',async()=>{
 const empty=await PDFDocument.create();empty.addPage();const bytes=await empty.save();
 await assert.rejects(compare(bytes,bytes),/没有可提取/);
 await assert.rejects(compare(oldBytes,new Uint8Array(12)),/有效的 PDF/);
 assert.equal(engine.pdf.sessions.size,0);
});

test('Unicode grapheme differences and span splitting do not corrupt line styles',()=>{
 const page={index:0,key:'page'},span=(text,x,index)=>({text,itemIndex:index,baseline:[x,50],angle:0,size:12,rect:[x,40,text.length*8,12],styles:Array.from(text,()=>({fontFamily:'CJK',fontSize:12,color:'#000000',bold:false,italic:false})).flatMap((s,i)=>Array.from({length:[...text][i].length},()=>s))});
 const a=groupSpans([span('预算金额：100万',20,0)],page),b=groupSpans([span('预算金额：',20,0),span('120万',60,1)],page);
 assert.equal(a[0].text,'预算金额：100万');assert.equal(b[0].text,'预算金额：120万');
 const model=lines=>({pages:[{kind:'page',index:0,text:lines[0].text,key:'page',label:'第一页',lines,geometry:{unit:'pt',space:'pageViewport',x:0,y:0,width:595,height:842,rotation:0},source:{locator:{type:'pdf',pageIndex:0,coordinateSpace:'pageViewport',rect:[0,0,595,842]}}}],nodes:[]});
 const c=comparePdfModels(model(a),model(b),differ).changes;assert.equal(c.filter(c=>c.aspect==='text').length,1);assert.equal(c.filter(c=>c.aspect==='style').length,0);
 assert.equal(differ.compare('金额😀100','金额😀120').edits[0].before.text,'0');
 const duplicate=matchObjects([{text:'相同标题'},{text:'相同标题'}],[{text:'相同标题'},{text:'相同标题'}],n=>n.text);assert.ok(duplicate.ambiguous.length);
});

test('real CJK PDF uses local CMaps and retains Chinese text in highlighted exports',async()=>{
 const a=await createCjkPdfFixture(),b=await createCjkPdfFixture(true),result=await compare(a,b);
 try{
  const text=result.changeSet.changes.find(c=>c.aspect==='text');assert.equal(text.before.value.data,'预算审阅：项目金额 120 万元');assert.equal(text.after.value.data,'预算审阅：项目金额 150 万元');
  assert.deepEqual(text.extensions['ppt2wiki.textDiff'].edits.map(e=>[e.before.text,e.after.text]),[['2','5']]);
  const out=await engine.annotate(b,result.changeSet,'after'),view=await engine.extract(out,'export.pdf');try{assert.equal(view.pages[0].text,'预算审阅：项目金额 150 万元');}finally{await engine.pdf.dispose(view);}
 }finally{await engine.disposeResult(result);}
});

test('repeated lines pair in bounded time and oversized fuzzy matching fails explicitly',()=>{
 const old=Array.from({length:500},()=>({text:'Repeated heading'})),next=Array.from({length:500},()=>({text:'Repeated heading'}));
 assert.equal(matchObjects(old,next,n=>n.text).pairs.length,500);
 assert.throws(()=>matchObjects(Array.from({length:400},(_,i)=>({text:'a'+i})),Array.from({length:400},(_,i)=>({text:'b'+i})),n=>n.text),/计算上限/);
});
