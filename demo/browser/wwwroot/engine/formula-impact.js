// Static reference analysis and saved-cache observations; this module does not recalculate formulas.
const sides=['before','after'];
const position=address=>{const m=/^([A-Z]+)(\d+)$/.exec(address);if(!m)return null;let column=0;for(const c of m[1])column=column*26+c.charCodeAt(0)-64;return {column,row:Number(m[2])};};
const address=raw=>raw.replaceAll('$','').toUpperCase();
const cell='\\$?[A-Za-z]{1,3}\\$?[1-9]\\d*';
const reference=new RegExp(`^(?:(?:'((?:[^']|'')+)'|([\\p{L}_\\\\][\\p{L}\\p{N}_.\\\\]*))!)?(${cell})(?::(${cell}))?`,'u');
const columnRange=/^(\$?[A-Za-z]{1,3}):(\$?[A-Za-z]{1,3})(?![\p{L}\p{N}_])/u;
const rowRange=/^(\$?[1-9]\d*):(\$?[1-9]\d*)(?![\p{L}\p{N}_])/u;
export function parseFormulaReferences(formula){
 const references=[],issues=[];let i=0;
 if(formula.length>32768)return {references,issues:['公式长度超过静态分析上限。']};
 while(i<formula.length){
  const rest=formula.slice(i),c=formula[i];
  if(c==='"'){i++;while(i<formula.length){if(formula[i++]==='"'){if(formula[i]==='"'){i++;continue;}break;}}continue;}
  if(c==='['){issues.push('外部或结构化引用尚未解析。');const end=formula.indexOf(']',i);i=end<0?formula.length:end+1;const external=reference.exec(formula.slice(i));if(external&&(external[1]||external[2]))i+=external[0].length;continue;}
  const m=reference.exec(rest);
  if(m&&!/[\p{L}\p{N}_.$!]/u.test(rest[m[0].length]||' ')&&rest[m[0].length]!=='('){
   if((m[1]||m[2]||'').includes('[')){issues.push('外部工作簿引用尚未解析。');i+=m[0].length;continue;}
   const start=address(m[3]),end=address(m[4]||m[3]),a=position(start),b=position(end);
   if(a.column<=16384&&b.column<=16384&&a.row<=1048576&&b.row<=1048576&&a.column<=b.column&&a.row<=b.row)references.push({sheet:m[1]?.replaceAll("''","'")||m[2]||null,start,end,expression:m[0]});
   else issues.push('引用区域越界或边界顺序不受支持。');
   i+=m[0].length;continue;
  }
  const columns=columnRange.exec(rest),rows=rowRange.exec(rest);
  if(columns||rows){const m=columns||rows;const start=columns?address(m[1])+'1':'A'+address(m[1]),end=columns?address(m[2])+'1048576':'XFD'+address(m[2]);const a=position(start),b=position(end);if(a.column<=b.column&&a.row<=b.row&&b.column<=16384&&b.row<=1048576)references.push({sheet:null,start,end,expression:m[0]});else issues.push('整行或整列引用越界。');i+=m[0].length;continue;}
  const name=/^[\p{L}_\\][\p{L}\p{N}_.\\]*/u.exec(rest);
  if(name){const token=name[0],next=rest.slice(token.length).trimStart();if(next.startsWith('(')){if(/^(?:_xlfn\.)?(INDIRECT|OFFSET)$/i.test(token))issues.push('动态引用的实际目标需要计算后确认。');else if(!/^(SUM|SUMIF|SUMIFS|AVERAGE|AVERAGEIF|MIN|MAX|COUNT|COUNTA|COUNTIF|COUNTIFS|IF|IFS|IFERROR|IFNA|ROUND|ROUNDUP|ROUNDDOWN|ABS|AND|OR|NOT|INT|MOD|PRODUCT|SUBTOTAL|SUMPRODUCT|VLOOKUP|HLOOKUP|INDEX|MATCH)$/i.test(token))issues.push('函数可能依赖未解析的输入或计算环境。');}else if(!/^(TRUE|FALSE)$/i.test(token))issues.push('命名、结构化或动态引用尚未完全解析。');i+=token.length;continue;}
  const numeric=/^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/.exec(rest);if(numeric){i+=numeric[0].length;continue;}
  if(c==='!'||c==='#'||c==='@')issues.push('引用表达式包含尚未解析的运算。');i++;
 }
 return {references,issues:[...new Set(issues)]};
}
const coverage=(issues=[])=>({status:issues.length?'partial':'complete',issues});
const issue=(code,detail,entityIds=[])=>({code,detail,entityIds});
export function buildFormulaImpacts({before,after,changeSet}){
 const views={before,after},entities=new Map(),dependencies=[],impacts=[],sheetEntities=new Map(),formulaIds=new Set(),edges={before:new Map(),after:new Map()};
 const changes=changeSet.changes,changesByEntity=new Map();for(const c of changes){if(!changesByEntity.has(c.target.entityId))changesByEntity.set(c.target.entityId,[]);changesByEntity.get(c.target.entityId).push(c);}
 const nodes={before:new Map(),after:new Map()},byAddress={before:new Map(),after:new Map()};
 for(const side of sides){for(const page of views[side].pages){const id='impact-'+page.key;sheetEntities.set(side+':'+page.key,id);if(!entities.has(id))entities.set(id,{entityId:id,kind:'sheet',label:page.name,before:null,after:null});entities.get(id)[side]={source:page.source,observations:[]};}
  for(const n of views[side].nodes){nodes[side].set(n.entityId,n);byAddress[side].set(n.pageKey+':'+n.source.locator.address,n);if(n.formula!==null)formulaIds.add(n.entityId);}
 }
 const register=id=>{if(entities.has(id))return entities.get(id);const n=nodes.after.get(id)||nodes.before.get(id);if(!n)return null;const e={entityId:id,kind:n.kind,label:n.label,before:null,after:null};for(const side of sides){const cell=nodes[side].get(id);if(!cell)continue;const cached=cell.formula!==null,v=cached?cell.cachedValue:cell.scalar;e[side]={source:cell.source,...(cell.formula!==''?{formula:cell.formula}:{}),observations:[{property:'/value',status:v?'known':'unavailable',origin:cached?'storedCache':'literal',freshness:cached||!v?'unverified':'notApplicable',...(v?{value:v}:{})}]};}entities.set(id,e);return e;};
 let scanBudget=250000;const dependencyIssues=[];
 const relevant={};for(const side of sides){relevant[side]=new Map();for(const n of views[side].nodes)if(formulaIds.has(n.entityId)||changesByEntity.has(n.entityId)){if(!relevant[side].has(n.pageKey))relevant[side].set(n.pageKey,[]);relevant[side].get(n.pageKey).push(n);}}
 for(const id of formulaIds){register(id);const refs={before:[],after:[]};
  for(const side of sides){const n=nodes[side].get(id);if(n?.formula==null)continue;const parsed=parseFormulaReferences(n.formula);if(n.formulaType!=='normal')parsed.issues.push('共享、数组或数据表公式尚未展开。');
   for(const ref of parsed.references){const page=ref.sheet?views[side].pages.find(p=>p.name.toLocaleLowerCase()===ref.sheet.toLocaleLowerCase()):views[side].pages.find(p=>p.key===n.pageKey);if(!page){parsed.issues.push('引用的工作表或外部输入无法定位。');continue;}
    const a=position(ref.start),b=position(ref.end),members=new Set();
    if(ref.start===ref.end){const input=byAddress[side].get(page.key+':'+ref.start);if(input){members.add(input.entityId);register(input.entityId);}}
    for(const input of relevant[side].get(page.key)||[]){if(--scanBudget<0){parsed.issues.push('引用传播分析已达到处理上限。');break;}if(input.row>=a.row&&input.row<=b.row&&input.column>=a.column&&input.column<=b.column){members.add(input.entityId);register(input.entityId);}}
    refs[side].push({kind:ref.start===ref.end?'cell':'range',expression:ref.expression,resolution:'resolved',members:[...members],areas:[{sheetEntityId:sheetEntities.get(side+':'+page.key),start:ref.start,end:ref.end}]});
   }
   if(parsed.issues.length){dependencyIssues.push(issue('unsupportedReference',parsed.issues.join(' '),[id]));for(const ref of refs[side])ref.resolution='partial';refs[side].push({kind:'dynamic',expression:n.formula||'未展开公式',resolution:'unresolved',members:[]});}
  }
  for(let i=0;i<Math.max(refs.before.length,refs.after.length);i++){const d={id:'dependency-'+(dependencies.length+1),dependentEntityId:id,before:refs.before[i]||null,after:refs.after[i]||null};dependencies.push(d);for(const side of sides)for(const member of d[side]?.members||[]){if(!edges[side].has(member))edges[side].set(member,[]);edges[side].get(member).push({id:d.id,target:id});}}
 }
 const causeMap=new Map();
 for(const c of changes.filter(c=>['value','formula','existence'].includes(c.aspect))){register(c.target.entityId);for(const side of sides){if(!c[side])continue;const queue=[{id:c.target.entityId,path:[]}],seen=new Set([c.target.entityId]);for(let head=0;head<queue.length;head++){const current=queue[head];if(current.path.length){if(!causeMap.has(current.id))causeMap.set(current.id,new Map());const key=c.id,previous=causeMap.get(current.id).get(key);if(!previous||previous.dependencyIds.length>current.path.length)causeMap.get(current.id).set(key,{kind:'change',changeId:c.id,dependencyIds:current.path});}
    for(const edge of edges[side].get(current.id)||[]){if(seen.has(edge.target))continue;seen.add(edge.target);queue.push({id:edge.target,path:[...current.path,edge.id]});}
   }}
 }
 for(const id of formulaIds){const own=changesByEntity.get(id)||[],valueChanges=own.filter(c=>c.aspect==='value'),formulaChanges=own.filter(c=>c.aspect==='formula');const causes=[...(causeMap.get(id)?.values()||[]),...formulaChanges.map(c=>({kind:'change',changeId:c.id,dependencyIds:[]}))].filter(c=>!valueChanges.some(v=>v.id===c.changeId));
  // Incomplete reference coverage alone is a limitation, not evidence of an impact.
  if(!valueChanges.length&&!formulaChanges.length&&!causes.length)continue;
  if(!causes.length)causes.push({kind:'unresolved',detail:'已保存结果存在差异，变化来源尚未确认。'});
  impacts.push({id:'impact-'+(impacts.length+1),mechanism:'formula',entityId:id,property:'/value',outcome:valueChanges.length?'changed':'unknown',changeIds:valueChanges.map(c=>c.id),causes,evidence:{level:valueChanges.length?'supported':'potential',method:valueChanges.length?'snapshotComparison':'staticAnalysis',detail:valueChanges.length?'比较的是文件保存的计算结果；引用关系来自静态解析，缓存可能未重算，原因尚未验证。':'关联输入或公式发生变化；相同或缺失的文件缓存不能确认重算结果。',computationIds:[]}});
 }
 const unknown=impacts.filter(i=>i.outcome==='unknown').map(i=>i.entityId);
 return {version:'1.0.0',scope:{entityIds:[...formulaIds],mechanisms:['formula']},entities:[...entities.values()],dependencies,computations:[],impacts,lineages:[],coverage:{dependencies:coverage(dependencyIssues),attribution:impacts.length?coverage([issue('unverifiedCause','引用来自静态解析，尚未进行因果回放验证。',impacts.map(i=>i.entityId))]):coverage(),computation:impacts.length?coverage([issue('unavailableResult','未运行计算引擎，文件缓存的新鲜度尚未验证。',unknown)]):{status:'notRequested',issues:[]},structure:{status:'notRequested',issues:[]}},delivery:{status:'complete',omitted:{entities:0,dependencies:0,computations:0,impacts:0,lineages:0}}};
}
