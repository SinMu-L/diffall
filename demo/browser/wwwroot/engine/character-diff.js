// Grapheme segmentation and UTF16 anchors stay in JS. Myers alignment runs in WASM.
export class CharacterDiffer {
  constructor(instance) { this.wasm = instance.exports; this.calls = 0; }
  static async load(bytes) {
    if (!bytes) {
      const response = await fetch(new URL('../character-diff.wasm', import.meta.url));
      if (!response.ok) throw new Error('无法加载字符差异 WASM 模块。');
      bytes = await response.arrayBuffer();
    }
    const {instance} = await WebAssembly.instantiate(bytes, {});
    return new CharacterDiffer(instance);
  }
  compare(before, after) {
    if (before.length + after.length > 2000000) throw new Error('单个文本对象超过字符比较范围（合计 200 万 UTF16 单元）。');
    const segmenter = new Intl.Segmenter(undefined, {granularity:'grapheme'});
    const a = [...segmenter.segment(before)], b = [...segmenter.segment(after)];
    const ids = new Map();
    const encode = tokens => Uint32Array.from(tokens, token => {
      if (!ids.has(token.segment)) ids.set(token.segment, ids.size + 1);
      return ids.get(token.segment);
    });
    const av = encode(a), bv = encode(b), count = a.length + b.length;
    const scratchLength = 2 * Math.ceil(count / 2) + 3;
    const required = (count + scratchLength * 2) * 4 + 64;
    const {memory, bisect} = this.wasm;
    if (memory.buffer.byteLength < required) memory.grow(Math.ceil((required - memory.buffer.byteLength) / 65536));
    new Uint32Array(memory.buffer, 0, av.length).set(av);
    new Uint32Array(memory.buffer, av.byteLength, bv.length).set(bv);
    const vf = count * 4, vr = vf + scratchLength * 4, matches = [];
    // Explicit task stack avoids JavaScript recursion limits on adversarial inputs.
    const tasks = [{os:0,oe:a.length,ns:0,ne:b.length}];
    while (tasks.length) {
      let {os,oe,ns,ne,match} = tasks.pop();
      if (match) { for (let i=0;i<oe-os;i++) matches.push([os+i,ns+i]); continue; }
      while(os<oe && ns<ne && av[os]===bv[ns]) matches.push([os++,ns++]);
      let suffix=0;
      while(os<oe-suffix && ns<ne-suffix && av[oe-suffix-1]===bv[ne-suffix-1]) suffix++;
      const oldEnd=oe-suffix,newEnd=ne-suffix,n=oldEnd-os,m=newEnd-ns;
      if(suffix) tasks.push({os:oldEnd,oe,ns:newEnd,ne,match:true});
      if(!n||!m) continue;
      if(n===1) { for(let j=ns;j<newEnd;j++) if(av[os]===bv[j]) {matches.push([os,j]);break;} continue; }
      if(m===1) { for(let i=os;i<oldEnd;i++) if(av[i]===bv[ns]) {matches.push([i,ns]);break;} continue; }
      const vocabulary=new Set(av.subarray(os,oldEnd));
      if(!bv.subarray(ns,newEnd).some(id=>vocabulary.has(id))) continue;
      this.calls++;
      const packed=bisect(os*4,n,av.byteLength+ns*4,m,vf,vr);
      if(packed<0n) throw new Error(packed===-2n?'文本差异计算量超出演示范围。':'WASM 字符对齐失败。');
      const x=Number(packed>>32n),y=Number(packed&0xffffffffn);
      if((!x&&!y)||(x===n&&y===m)) throw new Error('WASM 字符对齐未前进。');
      tasks.push({os:os+x,oe:oldEnd,ns:ns+y,ne:newEnd});
      tasks.push({os,oe:os+x,ns,ne:ns+y});
    }
    const edits=[];
    let ai=0,bi=0;
    const span=(text,tokens,start,end)=>{const s=tokens[start]?.index??text.length,e=tokens[end]?.index??text.length;return{start:s,length:e-s,text:text.slice(s,e)};};
    for(const [x,y] of [...matches,[a.length,b.length]]) {
      if(x!==ai||y!==bi) {
        const old=span(before,a,ai,x),next=span(after,b,bi,y);
        edits.push({operation:!old.length?'insert':!next.length?'delete':'replace',before:old,after:next});
      }
      ai=x+1;bi=y+1;
    }
    return {unit:'utf16',granularity:'grapheme',beforeLength:before.length,afterLength:after.length,edits};
  }
}
