// PDF coordinates use the unscaled, rotated page viewport (top left origin).
export const round=value=>Math.round(value*1000)/1000;
export const normalized=text=>text.replace(/\s+/gu,' ').trim();
const equal=(a,b)=>JSON.stringify(a)===JSON.stringify(b);
const value=data=>({type:typeof data==='string'?'text':typeof data==='number'?'number':Array.isArray(data)?'array':'object',data:typeof data==='number'?String(data):data});
export function similarity(a,b){
 a=normalized(a).slice(0,6000);b=normalized(b).slice(0,6000);if(a===b)return 1;if(!a||!b)return 0;
 const grams=s=>{const set=new Set();for(let i=0;i<s.length-1;i++)set.add(s.slice(i,i+2));return set;},x=grams(a),y=grams(b);
 return x.size+y.size?2*[...x].filter(t=>y.has(t)).length/(x.size+y.size):0;
}
/** Only unique exact anchors and unambiguous mutual best matches establish identity. */
export function matchObjects(old,next,text,threshold=.48){
 const a=new Set(old),b=new Set(next),pairs=[];
 const attach=(x,y,score,method)=>{pairs.push({old:x,next:y,score,method});a.delete(x);b.delete(y);};
 const oldKeys=new Map(),newKeys=new Map();
 for(const [items,keys]of [[old,oldKeys],[next,newKeys]])for(const item of items){const key=normalized(text(item));if(!key)continue;if(!keys.has(key))keys.set(key,[]);keys.get(key).push(item);}
 for(const [key,x]of oldKeys){const y=newKeys.get(key);if(x.length===1&&y?.length===1)attach(x[0],y[0],1,'content');}
 const ambiguous=[];
 if(threshold<=1)for(const [key,x]of oldKeys){const y=newKeys.get(key)||[];const remainingOld=x.filter(n=>a.has(n)),remainingNew=y.filter(n=>b.has(n));for(let i=0;i<Math.min(remainingOld.length,remainingNew.length);i++){attach(remainingOld[i],remainingNew[i],.65,'structure');ambiguous.push(remainingOld[i]);}}
 if(a.size*b.size>100000)throw new Error('PDF 待匹配内容过多，超过演示计算上限，请按章节拆分后比较。');
 const candidates=[],byOld=new Map(),byNew=new Map(),signatures=new Map();
 const signature=item=>{if(!signatures.has(item)){const s=normalized(text(item)).slice(0,6000),grams=new Set();for(let i=0;i<s.length-1;i++)grams.add(s.slice(i,i+2));signatures.set(item,{s,grams});}return signatures.get(item);};
 for(const x of a)for(const y of b){const sa=signature(x),sb=signature(y),score=sa.s===sb.s?1:sa.grams.size+sb.grams.size?2*[...sa.grams].filter(g=>sb.grams.has(g)).length/(sa.grams.size+sb.grams.size):0;
  if(score>=threshold){const candidate={x,y,score};candidates.push(candidate);for(const [map,key]of [[byOld,x],[byNew,y]]){if(!map.has(key))map.set(key,[]);map.get(key).push(candidate);}}
 }
 candidates.sort((x,y)=>y.score-x.score);
 for(const c of candidates){if(!a.has(c.x)||!b.has(c.y))continue;
  const rivals=[...(byOld.get(c.x)||[]),...(byNew.get(c.y)||[])].filter(v=>a.has(v.x)&&b.has(v.y)&&v!==c);
  if(rivals.some(v=>v.score>=c.score-.065)){ambiguous.push(c.x);continue;}
  attach(c.x,c.y,Math.min(c.score,.95),'composite');
 }
 // Empty pages carry no text identity; only retain their order with a coverage warning.
 if(threshold<=1)for(const x of [...a].filter(n=>!normalized(text(n)))){const y=[...b].find(n=>!normalized(text(n)));if(y){attach(x,y,.5,'structure');ambiguous.push(x);}}
 return {pairs,removed:[...a],added:[...b],ambiguous:[...new Set(ambiguous)]};
}
export function orderMoves(pairs){
 const old=[...pairs].sort((a,b)=>a.old.index-b.old.index),next=[...pairs].sort((a,b)=>a.next.index-b.next.index);
 const dp=Array.from({length:old.length+1},()=>new Float64Array(next.length+1));
 // Among equally long subsequences prefer pages that retain their actual page index.
 const weight=p=>1+(p.old.index===p.next.index?1/(pairs.length+1):0);
 for(let i=old.length-1;i>=0;i--)for(let j=next.length-1;j>=0;j--)dp[i][j]=old[i]===next[j]?dp[i+1][j+1]+weight(old[i]):Math.max(dp[i+1][j],dp[i][j+1]);
 const stable=new Set();let i=0,j=0;while(i<old.length&&j<next.length){if(old[i]===next[j]){stable.add(old[i]);i++;j++;}else if(dp[i+1][j]>=dp[i][j+1])i++;else j++;}
 return pairs.filter(p=>!stable.has(p)&&p.old.index!==p.next.index);
}
function ranges(chars,key){const result=[];for(let i=0;i<chars.length;i++){const v=chars[i]?.[key]??null,last=result.at(-1);if(last&&equal(last.value,v))last.length++;else result.push({start:i,length:1,value:v});}return result;}
function styleChanged(a,b,key,diff){
 let x=0,y=0;
 const compare=(length)=>{for(let i=0;i<length;i++){const av=a.styles[x+i]?.[key],bv=b.styles[y+i]?.[key];if(av!=null&&bv!=null&&!equal(av,bv))return true;}return false;};
 for(const edit of diff.edits){if(compare(edit.before.start-x))return true;x=edit.before.start+edit.before.length;y=edit.after.start+edit.after.length;}
 return compare(a.text.length-x);
}
export function comparePdfModels(before,after,differ){
 const pageMatches=matchObjects(before.pages,after.pages,p=>p.text,.6),changes=[],issues=[];let sequence=0;
 const identify=(a,b)=>{const id='pdf-entity-'+(++sequence);if(a)a.entityId=id;if(b)b.entityId=id;return id;};
 const target=n=>({entityId:n.entityId,kind:n.kind,label:n.label});
 const state=(n,data)=>({source:n.source,value:value(data)});
 const match=p=>({method:p.method,confidence:p.score,evidence:[p.score===1?'提取文字一致且在作用域内唯一':p.method==='structure'?'重复文字按提取顺序对应，身份存在不确定性':'双向最佳文字匹配，且与其他候选存在分差']});
 const add=c=>changes.push({id:'change-'+String(changes.length+1).padStart(3,'0'),...c});
 const update=(p,aspect,property,a,b,extension)=>{if(equal(a,b))return;add({operation:'update',aspect,property,target:target(p.next),before:state(p.old,a),after:state(p.next,b),match:match(p),...(extension?{extensions:extension}:{})});};
 const existence=(n,operation)=>{const s=state(n,{kind:n.kind,text:n.text});add({operation,aspect:'existence',property:'',target:target(n),before:operation==='remove'?s:null,after:operation==='add'?s:null});};
 const matchedLines=[];
 for(const p of pageMatches.pairs){const id=identify(p.old,p.next),key='pdf:'+id;p.old.key=p.next.key=key;p.old.name=p.next.name=`旧 ${p.old.index+1} ↔ 新 ${p.next.index+1}`;
  for(const page of [p.old,p.next])for(const line of page.lines)line.pageKey=key;
  update(p,'geometry','/geometry',p.old.geometry,p.next.geometry);
  const lines=matchObjects(p.old.lines,p.next.lines,n=>n.text,.5);
  matchedLines.push(...lines.pairs);
  if(lines.ambiguous.length)issues.push({code:'ambiguousMatch',side:'both',message:`旧第 ${p.old.index+1} 页含重复或相似文字，部分对象身份无法唯一确定。`});
  p.removed=lines.removed;p.added=lines.added;
 }
 // Match unique unchanged text across matched pages, e.g. a paragraph flowing to a new page.
 const cross=matchObjects(pageMatches.pairs.flatMap(p=>p.removed),pageMatches.pairs.flatMap(p=>p.added),n=>n.text,1.01);
 matchedLines.push(...cross.pairs);
 for(const p of matchedLines){identify(p.old,p.next);const diff=differ.compare(p.old.text,p.next.text);
  update(p,'text','/text',p.old.text,p.next.text,{'ppt2wiki.textDiff':diff});
  for(const key of ['fontFamily','fontSize','color','bold','italic'])if(styleChanged(p.old,p.next,key,diff))update(p,'style','/style/'+key,ranges(p.old.styles,key),ranges(p.next.styles,key));
  // A content/size change naturally alters the box. Position and rotation remain separate facts.
  const a=p.old.geometry,b=p.next.geometry;
  if(Math.abs(a.x-b.x)>.75||Math.abs(a.y-b.y)>.75||Math.abs(a.rotation-b.rotation)>.1){
   // Font ascent changes are not a user movement: compare the baseline as well.
   if(Math.abs(p.old.baseline[0]-p.next.baseline[0])>.75||Math.abs(p.old.baseline[1]-p.next.baseline[1])>.75||Math.abs(a.rotation-b.rotation)>.1)update(p,'geometry','/geometry',a,b);
  }
  if(p.old.pageKey!==p.next.pageKey)add({operation:'move',aspect:'parent',property:'/placement',target:target(p.next),before:state(p.old,{parentEntityId:p.old.pageKey.slice(4),index:p.old.index}),after:state(p.next,{parentEntityId:p.next.pageKey.slice(4),index:p.next.index}),match:match(p)});
 }
 for(const n of cross.removed){identify(n);existence(n,'remove');}for(const n of cross.added){identify(null,n);existence(n,'add');}
 for(const p of orderMoves(pageMatches.pairs))add({operation:'move',aspect:'order',property:'/placement',target:target(p.next),before:state(p.old,{parentEntityId:'document',index:p.old.index}),after:state(p.next,{parentEntityId:'document',index:p.next.index}),match:match(p)});
 for(const [pages,operation]of [[pageMatches.removed,'remove'],[pageMatches.added,'add']])for(const page of pages){identify(operation==='remove'?page:null,operation==='add'?page:null);page.key='pdf:'+page.entityId;page.name=(operation==='remove'?'删除 · 旧':'新增 · 新')+'第 '+(page.index+1)+' 页';for(const line of page.lines)line.pageKey=page.key;existence(page,operation);}
 if(pageMatches.ambiguous.length)issues.push({code:'ambiguousMatch',side:'both',message:'部分页面文字相似或重复，页面对应关系存在不确定性。'});
 for(const view of [before,after])view.nodes=view.pages.flatMap(p=>[p,...p.lines]);
 return {changes,issues};
}

/** Stable line grouping removes parser run boundaries without joining distant columns. */
export function groupSpans(spans,page){
 const lines=[];
 for(const span of [...spans].sort((a,b)=>Math.abs(a.baseline[1]-b.baseline[1])>2?a.baseline[1]-b.baseline[1]:a.rect[0]-b.rect[0])){
  const last=lines.at(-1),gap=last?span.rect[0]-last.rect[0]-last.rect[2]:Infinity;
  if(last&&Math.abs(last.baseline[1]-span.baseline[1])<2&&Math.abs(last.angle-span.angle)<.1&&gap>=-2&&gap<Math.max(15,span.size*1.8)){
   const space=gap>span.size*.18&&!/\s$/.test(last.text)&&!/^\s/.test(span.text)?' ':'';
   last.text+=space+span.text;if(space)last.styles.push(last.styles.at(-1));last.styles.push(...span.styles);last.itemIndices.push(span.itemIndex);
   const x=Math.min(last.rect[0],span.rect[0]),y=Math.min(last.rect[1],span.rect[1]);last.rect=[x,y,Math.max(last.rect[0]+last.rect[2],span.rect[0]+span.rect[2])-x,Math.max(last.rect[1]+last.rect[3],span.rect[1]+span.rect[3])-y];
  }else lines.push({...span,itemIndices:[span.itemIndex]});
 }
 return lines.map((line,index)=>{const rect=line.rect.map(round);return {kind:'textRange',index,key:`pdf:${page.index}:line:${index}`,pageKey:page.key,label:`第 ${page.index+1} 页 · ${line.text.slice(0,35)}`,text:line.text,styles:line.styles,baseline:line.baseline.map(round),geometry:{unit:'pt',space:'pageViewport',x:rect[0],y:rect[1],width:rect[2],height:rect[3],rotation:round(line.angle)},source:{locator:{type:'pdf',pageIndex:page.index,coordinateSpace:'pageViewport',rect,itemIndices:line.itemIndices,text:line.text,range:{unit:'utf16',start:0,length:line.text.length}}}};});
}
