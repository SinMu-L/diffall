import test from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import {renderCharacterText,characterDiff} from '../browser/wwwroot/text-diff-view.js';
import {CharacterDiffer} from '../browser/wwwroot/engine/character-diff.js';
const differ=await CharacterDiffer.load(readFileSync(new URL('../browser/wwwroot/character-diff.wasm',import.meta.url)));
const diff={unit:'utf16',granularity:'grapheme',beforeLength:6,afterLength:6,edits:[{operation:'replace',before:{start:1,length:1,text:'2'},after:{start:1,length:1,text:'5'}}]};
test('only the changed digit is marked and context remains plain',()=>{
 assert.equal(renderCharacterText('120 万元',diff,'before'),'1<del class="char-diff before" data-start="1" data-length="1">2</del>0 万元');
 assert.equal(renderCharacterText('150 万元',diff,'after'),'1<ins class="char-diff after" data-start="1" data-length="1">5</ins>0 万元');
});
test('formatted numbers keep commas and mark the corresponding digit',()=>{
 assert.equal(renderCharacterText('1,200,000',diff,'before',{rawText:'1200000'}),'1,<del class="char-diff before" data-start="1" data-length="1">2</del>00,000');
});
test('numeric preview preserves significant digits instead of hiding changes by rounding',()=>{
 const value={edits:[{before:{start:9,length:1,text:'6'},after:{start:9,length:1,text:'7'}}]};
 assert.match(renderCharacterText('300.123',value,'before',{rawText:'300.123456'}),/>6<\/del>/);
});
test('turning highlights off restores the unmarked text',()=>{
 assert.equal(renderCharacterText('120 万元',diff,'before',{enabled:false}),'120 万元');
});
test('characters and context are HTML escaped',()=>{
 const unsafe={edits:[{before:{start:0,length:1,text:'<'},after:{start:0,length:1,text:'>'}}]};
 const html=renderCharacterText('<script>',unsafe,'before');assert.ok(html.includes('&lt;</del>script&gt;'));assert.ok(!html.includes('<script>'));
});
test('insertion and deletion anchors do not inject replacement text into the other version',()=>{
 const insertion={edits:[{before:{start:1,length:0,text:''},after:{start:1,length:1,text:'新'}}]};
 assert.match(renderCharacterText('甲乙',insertion,'before'),/甲<span class="char-caret before"[^>]+><\/span>乙/);
 assert.match(renderCharacterText('甲新乙',insertion,'after'),/>新<\/ins>/);
});
test('mismatched source spans fail instead of marking an unrelated character',()=>assert.throws(()=>renderCharacterText('190 万元',diff,'before'),/不一致/));
test('real engine payloads highlight two separated Word edits and a single PPT digit',()=>{
 const facts=JSON.parse(readFileSync(new URL('./fixtures/change-facts.json',import.meta.url),'utf8'));
 const load=name=>{const change=facts[name].find(change=>change.aspect==='text');return{...change,extensions:{'ppt2wiki.textDiff':differ.compare(change.before.value.data,change.after.value.data)}};};
 const word=load('word-budget'),ppt=load('ppt-budget');
 assert.equal(characterDiff(word).edits.length,2);
 assert.equal((renderCharacterText(word.before.value.data,characterDiff(word),'before').match(/<del /g)||[]).length,2);
 assert.equal(characterDiff(ppt).edits[0].before.text,'2');assert.equal(characterDiff(ppt).edits[0].after.text,'5');
});
