import test from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync,writeFileSync,mkdirSync} from 'node:fs';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {zipSync,strToU8} from 'fflate';
globalThis.DOMParser=DOMParser;globalThis.XMLSerializer=XMLSerializer;
const {OfficeEngine,normalizeNumber}=await import('../browser/wwwroot/engine/office-engine.js');
const {CharacterDiffer}=await import('../browser/wwwroot/engine/character-diff.js');
const {OfficePackage,NS,sha256,wordText,drawingText,descendants,first,elements,children,attr,nsAttr,create,append,remove}=await import('../browser/wwwroot/engine/ooxml.js');
const differ=await CharacterDiffer.load(readFileSync(new URL('../browser/wwwroot/character-diff.wasm',import.meta.url))),engine=new OfficeEngine(differ);
const read=name=>new Uint8Array(readFileSync(new URL('../browser/wwwroot/samples/'+name,import.meta.url)));
const cases=[['word-budget','docx',3],['excel-budget','xlsx',13],['ppt-budget','pptx',7]];
const expectedFacts=JSON.parse(readFileSync(new URL('./fixtures/change-facts.json',import.meta.url),'utf8'));
mkdirSync(new URL('../artifacts/',import.meta.url),{recursive:true});
for(const [name,ext,count]of cases)test(`${ext}: real files, Change protocol and both native annotation sides`,async()=>{
  const oldName=name+'-old.'+ext,newName=name+'-new.'+ext,before=read(oldName),after=read(newName);
  const result=await engine.compare(before,after,oldName,newName);assert.equal(result.changeSet.changes.length,count);
  const facts=set=>set.changes.map(c=>({operation:c.operation,aspect:c.aspect,property:c.property,before:c.before,after:c.after}));
  assert.deepEqual(facts(result.changeSet),expectedFacts[name]);
  writeFileSync(new URL('../artifacts/'+name+'.changes.json',import.meta.url),JSON.stringify(result.changeSet,null,2));
  assert.equal((await engine.compare(before,before,oldName,oldName)).changeSet.changes.length,0);
  for(const [side,bytes,filename]of [['before',before,oldName],['after',after,newName]]){
    const output=await engine.annotate(bytes,result.changeSet,side),sourcePkg=new OfficePackage(bytes),pkg=new OfficePackage(output);
    writeFileSync(new URL(`../artifacts/${name}.${side}.frontend.${ext}`,import.meta.url),output);
    const original=await engine.extract(bytes,filename),annotated=await engine.extract(output,filename),bySource=new Map(annotated.nodes.map(n=>[JSON.stringify(n.source),n]));
    for(const n of original.nodes){const found=bySource.get(JSON.stringify(n.source));assert.ok(found);for(const key of ['text','scalar','formula','display'])assert.deepEqual(found[key],n[key],key+' preserved');}
    // Untouched binary assets and parts retain exact payload bytes.
    const touched=new Set(result.changeSet.changes.flatMap(c=>c[side]?[c[side].source.part.slice(1)]:[]));
    for(const [part,data]of Object.entries(sourcePkg.files))if(!touched.has(part)&&!part.endsWith('styles.xml'))assert.deepEqual(pkg.files[part],data,part);
    const textChanges=result.changeSet.changes.filter(c=>['text','value'].includes(c.aspect)&&c[side]?.value.type==='text'&&c.extensions?.['ppt2wiki.textDiff']);
    for(const change of textChanges){const spans=change.extensions['ppt2wiki.textDiff'].edits.map(edit=>edit[side]).filter(s=>s.length),expected=spans.map(s=>s.text).join('');const root=pkg.xml(change[side].source.part);
      if(ext==='docx'){const p=descendants(root,'w','p').find(p=>wordText(p)===change[side].value.data);assert.ok(p);const marked=descendants(p,'w','r').filter(r=>first(first(r,'w','rPr'),'w','shd')).map(r=>children(r).filter(c=>c.localName==='t').map(c=>c.textContent).join('')).join('');assert.equal(marked,expected);assert.equal(first(first(p,'w','pPr'),'w','shd'),undefined);}
      if(ext==='xlsx'){const cell=descendants(root,'s','c').find(c=>attr(c,'r')===change[side].source.locator.address);const marked=descendants(cell,'s','r').filter(r=>first(first(r,'s','rPr'),'s','u')).map(r=>first(r,'s','t').textContent).join('');assert.equal(marked,expected);}
      if(ext==='pptx'){const loc=change[side].source.locator,native=descendants(root,'p','cNvPr').find(n=>Number(attr(n,'id'))===loc.shapePath.at(-1));let shape=native.parentNode.parentNode,body;if(loc.cell){body=first(elements(elements(descendants(shape,'a','tbl')[0],'a','tr')[loc.cell.row],'a','tc')[loc.cell.column],'a','txBody');}else body=first(shape,'p','txBody');const marked=descendants(body,'a','r').filter(r=>first(first(r,'a','rPr'),'a','highlight')).map(r=>first(r,'a','t').textContent).join('');assert.equal(marked,expected);}
    }
  }
  await assert.rejects(engine.annotate(after,result.changeSet,'before'),/快照/);
});
test('decimal normalization never loses significant digits',()=>{assert.equal(normalizeNumber('1200000'),normalizeNumber('1.200e6'));assert.notEqual(normalizeNumber('9007199254740992'),normalizeNumber('9007199254740993'));});
test('Word split runs retain bold and italic on each original character',async()=>{
  const old=read('word-budget-old.docx'),original=new OfficePackage(old),part=original.main();
  function customize(texts){const pkg=new OfficePackage(old),p=descendants(pkg.xml(part),'w','p')[4];for(const c of [...children(p)])if(c.localName!=='pPr')remove(c);texts.forEach((text,i)=>{const run=append(p,'w','r'),pr=append(run,'w','rPr');append(pr,'w',i%2?'b':'i');append(run,'w','t',{},text);});pkg.mark(part);return pkg.write();}
  const a=customize(['AA旧','预算','BB']),b=customize(['AA新','计划','BB']);const result=await engine.compare(a,b,'a.docx','b.docx');const annotated=await engine.annotate(b,result.changeSet,'after'),pkg=new OfficePackage(annotated),p=descendants(pkg.xml(part),'w','p')[4];assert.equal(wordText(p),'AA新计划BB');const signature=node=>descendants(node,'w','r').flatMap(run=>[...first(run,'w','t').textContent].map(char=>[char,!!first(first(run,'w','rPr'),'w','b'),!!first(first(run,'w','rPr'),'w','i')]));assert.deepEqual(signature(p),signature(descendants(new OfficePackage(b).xml(part),'w','p')[4]));
});
test('Excel shared-string annotation isolates the changed cell',async()=>{
  const pkg=new OfficePackage(read('excel-budget-old.xlsx')),main=pkg.main(),sheet=pkg.related(main,nsAttr(descendants(pkg.xml(main),'s','sheet')[0],'r','id'));
  const sstPart='xl/sharedStrings.xml';pkg.files[sstPart]=strToU8('<sst xmlns="'+NS.s+'" count="2" uniqueCount="1"><si><r><rPr><b/></rPr><t>李敏</t></r></si></sst>');
  const rel='xl/_rels/workbook.xml.rels';append(pkg.xml(rel),'rel','Relationship',{Id:'rIdTestSst',Type:NS.r+'/sharedStrings',Target:'sharedStrings.xml'});pkg.mark(rel);append(pkg.xml('[Content_Types].xml'),'ct','Override',{PartName:'/'+sstPart,ContentType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'});pkg.mark('[Content_Types].xml');
  for(const address of ['E2','E3']){const cell=descendants(pkg.xml(sheet),'s','c').find(c=>attr(c,'r')===address);for(const child of children(cell))remove(child);cell.setAttribute('t','s');append(cell,'s','v',{},'0');}pkg.mark(sheet);const input=pkg.write(),set={format:'xlsx',documents:{after:{sha256:await sha256(input)}},changes:[{aspect:'value',after:{source:{part:'/'+sheet,locator:{type:'excel',sheetId:1,address:'E2'}},value:{type:'text',data:'李敏'}},extensions:{'ppt2wiki.textDiff':{edits:[{after:{start:1,length:1,text:'敏'}}]}}}]};
  const out=new OfficePackage(await engine.annotate(input,set,'after'));assert.deepEqual(out.files[sstPart],pkg.files[sstPart]);assert.equal(attr(descendants(out.xml(sheet),'s','c').find(c=>attr(c,'r')==='E3'),'t'),'s');assert.equal(attr(descendants(out.xml(sheet),'s','c').find(c=>attr(c,'r')==='E2'),'t'),'inlineStr');
});
test('invalid archives, XML entities and mismatched formats fail clearly',async()=>{assert.throws(()=>new OfficePackage(new Uint8Array(12)),/ZIP/);const unsafe=zipSync({'[Content_Types].xml':strToU8('<!DOCTYPE x [<!ENTITY a "bad">]><Types/>')});assert.throws(()=>new OfficePackage(unsafe).xml('[Content_Types].xml'),/DTD/);await assert.rejects(engine.compare(read('word-budget-old.docx'),read('excel-budget-new.xlsx'),'a.docx','b.xlsx'),/同一种/);});

test('comparison waits for progress feedback and stops at an unreadable revision',async()=>{
  const stages=[];let resume;const paint=new Promise(resolve=>{resume=resolve;});
  const pending=engine.compare(read('word-budget-old.docx'),new Uint8Array(12),'old.docx','new.docx',{onProgress:stage=>{stages.push(stage);if(stage==='before')return paint;}});
  let settled=false;const rejected=assert.rejects(pending,/ZIP/).then(()=>{settled=true;});
  await new Promise(resolve=>setImmediate(resolve));
  assert.equal(settled,false,'parsing must wait until the UI has painted its progress');assert.deepEqual(stages,['before']);
  resume();await rejected;assert.deepEqual(stages,['before','after'],'failure must not report a later computation stage');
  stages.length=0;const bytes=read('word-budget-old.docx');
  const result=await engine.compare(bytes,bytes,'old.docx','new.docx',{onProgress:stage=>stages.push(stage)});
  assert.deepEqual(stages,['before','after','diff']);assert.equal(result.changeSet.changes.length,0);
});
