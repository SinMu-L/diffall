import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import {webcrypto} from 'node:crypto';
import {readFile,writeFile,mkdir} from 'node:fs/promises';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {zipSync,strToU8,unzipSync} from 'fflate';
import {OfficeEngine} from '../browser/wwwroot/engine/office-engine.js';
import {CharacterDiffer} from '../browser/wwwroot/engine/character-diff.js';

globalThis.DOMParser=DOMParser;globalThis.XMLSerializer=XMLSerializer;
const source=await readFile(new URL('../../skills/office-diff/assets/office-diff.bundle.js',import.meta.url),'utf8');
const context=vm.createContext({console,TextEncoder,TextDecoder,Uint8Array,Intl,crypto:webcrypto,atob,btoa,fetch(){throw new Error('Unexpected network request');}});
const api=vm.runInContext(source,context,{timeout:10000});
const original=new OfficeEngine(await CharacterDiffer.load(await readFile(new URL('../browser/wwwroot/character-diff.wasm',import.meta.url))));
const fixtures=new URL('../browser/wwwroot/samples/',import.meta.url);
const output=new URL('../artifacts/skill-bundle/',import.meta.url);
await mkdir(output,{recursive:true});
const b64=bytes=>Buffer.from(bytes).toString('base64');

test('self-contained bundle works without browser DOM, process, require or fetch',async()=>{
  assert.equal(JSON.parse(await api.selfCheck()).ok,true);
  assert.equal(context.DOMParser,undefined);
  assert.equal(context.XMLSerializer,undefined);
  assert.equal(context.process,undefined);
  assert.equal(context.require,undefined);
});

for(const [name,format,count] of [['word-budget','docx',3],['excel-budget','xlsx',13],['ppt-budget','pptx',7],['excel-formulas','xlsx',null]]){
  test(`${format} ${name}: bundled engine matches every original ChangeSet field`,async()=>{
    const oldName=name+'-old.'+format,newName=name+'-new.'+format;
    const before=await readFile(new URL(oldName,fixtures)),after=await readFile(new URL(newName,fixtures));
    const expected=(await original.compare(before,after,oldName,newName)).changeSet;
    const actual=JSON.parse(await api.compareBase64(b64(before),b64(after),oldName,newName));
    assert.deepEqual(actual,expected);
    if(count!==null)assert.equal(actual.changes.length,count);
    assert.equal(JSON.parse(await api.compareBase64(b64(before),b64(before),oldName,oldName)).changes.length,0);
    assert.deepEqual(await readFile(new URL(oldName,fixtures)),before);
    await writeFile(new URL(name+'.json',output),JSON.stringify(actual,null,2));
  });
}

test('annotation remains snapshot-bound in the bundled engine',async()=>{
  const before=await readFile(new URL('word-budget-old.docx',fixtures)),after=await readFile(new URL('word-budget-new.docx',fixtures));
  const set=await api.compareBase64(b64(before),b64(after),'old.docx','new.docx');
  const annotated=Buffer.from(await api.annotateBase64(b64(after),set,'after'),'base64');
  assert.ok(unzipSync(annotated)['word/document.xml']);
  await assert.rejects(api.annotateBase64(b64(before),set,'after'),/快照/);
});

test('wrong formats, malformed XML and DTD fail instead of reporting zero changes',async()=>{
  await assert.rejects(api.compareBase64('','','a.pdf','b.pdf'),/不包含 PDF/);
  const unsafe=zipSync({'[Content_Types].xml':strToU8('<!DOCTYPE x [<!ENTITY a "bad">]><Types/>'),'_rels/.rels':strToU8('<Relationships/>')});
  const bad=zipSync({'[Content_Types].xml':strToU8('<Types>'),'_rels/.rels':strToU8('<Relationships>')});
  await assert.rejects(api.compareBase64(b64(unsafe),b64(unsafe),'a.docx','b.docx'));
  await assert.rejects(api.compareBase64(b64(bad),b64(bad),'a.docx','b.docx'));
});

test('delivery archive contains the full runnable folder and matching resource hashes',async()=>{
  const files=unzipSync(await readFile(new URL('../../artifacts/office-diff-skill.zip',import.meta.url)));
  for(const name of ['SKILL.md','scripts/office_diff.py','assets/office-diff.bundle.js','assets/engine-manifest.json','assets/THIRD_PARTY_LICENSES.txt','references/runtime.md','references/protocol.md'])assert.ok(files['office-diff/'+name],name);
  assert.equal(new TextDecoder().decode(files['office-diff/assets/office-diff.bundle.js']),source);
  assert.ok(Object.keys(files).every(name=>!name.includes('__pycache__')&&!name.includes('..')));
});
