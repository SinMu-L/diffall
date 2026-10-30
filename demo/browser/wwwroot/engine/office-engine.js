import {OfficePackage,sha256,is} from './ooxml.js';
import {extractWord,extractExcel,extractPowerPoint} from './adapters.js';
import {annotate} from './annotation.js';
import {buildFormulaImpacts} from './formula-impact.js';
import {PdfEngine} from './pdf-engine.js';
const V=(type,data)=>({type,data});
const equal=(a,b)=>JSON.stringify(a)===JSON.stringify(b);
export function normalizeNumber(value){const m=/^(-?)(\d+)(?:\.(\d+))?(?:[eE]([+-]?\d+))?$/.exec(value);if(!m)return value;let digits=(m[2]+(m[3]??'')).replace(/^0+/,'');if(!digits)return'0';let exponent=BigInt(m[4]??0)-BigInt(m[3]?.length??0);const trimmed=digits.replace(/0+$/,'');exponent+=BigInt(digits.length-trimmed.length);return m[1]+trimmed+'e'+exponent;}
function equivalent(a,b){if(a.type!==b.type)return false;if(a.type==='asset')return a.data.sha256===b.data.sha256;if(a.type==='number')return normalizeNumber(a.data)===normalizeNumber(b.data);return equal(a,b);}
function similarity(a,b){if(a===b)return 1;if(!a||!b)return 0;const chunks=text=>new Set(text.match(/[\s\S]{1,2}/g)),x=chunks(a),y=chunks(b);return 2*[...x].filter(t=>y.has(t)).length/(x.size+y.size);}
function movedSlides(pairs){const slides=pairs.filter(p=>p.old.kind==='slide'),old=[...slides].sort((a,b)=>a.old.index-b.old.index),next=[...slides].sort((a,b)=>a.next.index-b.next.index),dp=Array.from({length:old.length+1},()=>new Uint32Array(next.length+1));for(let i=old.length-1;i>=0;i--)for(let j=next.length-1;j>=0;j--)dp[i][j]=old[i]===next[j]?1+dp[i+1][j+1]:Math.max(dp[i+1][j],dp[i][j+1]);let x=0,y=0;const stable=new Set();while(x<old.length&&y<next.length){if(old[x]===next[y]){stable.add(old[x]);x++;y++;}else if(dp[x+1][y]>=dp[x][y+1])x++;else y++;}return slides.filter(p=>!stable.has(p));}
export function semanticDiff(before,after,differ,{compareCachedFormulaValues=false}={}){
  const old=[...before.nodes],next=[...after.nodes],pairs=[];
  for(const a of [...old]){const candidates=next.filter(b=>a.kind===b.kind&&a.key===b.key&&a.strongKey&&b.strongKey);if(candidates.length!==1||old.filter(b=>a.kind===b.kind&&a.key===b.key).length!==1)continue;const b=candidates[0];pairs.push({old:a,next:b,score:.98,method:'composite'});old.splice(old.indexOf(a),1);next.splice(next.indexOf(b),1);}
  for(const a of old.filter(n=>n.kind==='paragraph')){const scored=next.filter(b=>b.kind===a.kind).map(b=>({node:b,score:similarity(a.text??'',b.text??'')})).sort((a,b)=>b.score-a.score);if(!scored.length||scored[0].score<.62||(scored[1]&&scored[0].score-scored[1].score<.12))continue;const best=scored[0];pairs.push({old:a,next:best.node,score:best.score,method:'content'});old.splice(old.indexOf(a),1);next.splice(next.indexOf(best.node),1);}
  let sequence=0;const entity=new Map();for(const pair of pairs){const id='entity-'+(++sequence);entity.set(pair.old,id);entity.set(pair.next,id);}for(const n of [...old,...next])entity.set(n,'entity-'+(++sequence));
  for(const [n,entityId] of entity)n.entityId=entityId;
  const changes=[],target=n=>({entityId:entity.get(n),kind:n.kind,label:n.label}),state=(n,value)=>({value,source:n.source}),match=p=>({method:p.method,confidence:p.score,evidence:[p.method==='content'?'段落内容相似且候选匹配唯一':'格式适配器作用域内的对象标识或行列键一致','对象类型一致']}),id=()=> 'change-'+String(changes.length+1).padStart(3,'0');
  function update(pair,aspect,property,a,b){if(equivalent(a,b))return;const change={id:id(),operation:'update',aspect,target:target(pair.next),property,before:state(pair.old,a),after:state(pair.next,b),match:match(pair)};if(['text','value','formula'].includes(aspect)&&a.type==='text'&&b.type==='text')change.extensions={'ppt2wiki.textDiff':differ.compare(a.data,b.data)};changes.push(change);}
  for(const pair of pairs){const a=pair.old,b=pair.next;if(a.kind==='slide')continue;
    if(a.text!==null&&b.text!==null)update(pair,'text','/text',V('text',a.text),V('text',b.text));
    if(a.formula!==null||b.formula!==null){
      const av=a.formula!==null?a.cachedValue:a.scalar,bv=b.formula!==null?b.cachedValue:b.scalar;
      if(compareCachedFormulaValues&&av&&bv)update(pair,'value','/value',av,bv);
    }else if(a.scalar&&b.scalar)update(pair,'value','/value',a.scalar,b.scalar);
    if((a.formula!==null||b.formula!==null)&&a.formula!==''&&b.formula!=='')update(pair,'formula','/formula',V(a.formula===null?'null':'text',a.formula),V(b.formula===null?'null':'text',b.formula));
    if(['paragraph','shape','cell'].includes(a.kind)){update(pair,'style','/style/font/bold',V('boolean',a.bold),V('boolean',b.bold));update(pair,'style','/style/font/size',V('number',String(a.fontSize)),V('number',String(b.fontSize)));update(pair,'style','/style/font/color',V('text',a.color),V('text',b.color));}
    if(a.geometry&&b.geometry)update(pair,'geometry','/geometry',V('object',a.geometry),V('object',b.geometry));
    if(a.asset&&b.asset)update(pair,'media','/asset',V('asset',a.asset),V('asset',b.asset));
  }
  for(const pair of movedSlides(pairs))changes.push({id:id(),operation:'move',aspect:'order',target:target(pair.next),property:'/placement',before:state(pair.old,V('object',{parentEntityId:'document',index:pair.old.index})),after:state(pair.next,V('object',{parentEntityId:'document',index:pair.next.index})),match:match(pair)});
  for(const [nodes,operation]of [[old,'remove'],[next,'add']]){const missing=new Set(nodes.map(n=>n.key));for(const n of nodes.filter(n=>!missing.has(n.parentKey))){const s=state(n,V('object',{kind:n.kind,text:n.text,scalar:n.scalar}));changes.push({id:id(),operation,aspect:'existence',target:target(n),property:'',before:operation==='remove'?s:null,after:operation==='add'?s:null});}}
  return changes;
}
const limitations={docx:'当前比较正文段落文字和首个文字片段的粗体、字号、颜色；混合片段样式、页眉页脚、复杂字段、表格结构、主题继承和页面排版未完整比较。',xlsx:'当前比较工作表中的单元格值、普通公式和直接文字样式；合并、共享公式、主题继承、图表与行列结构未完整比较。',pptx:'当前比较幻灯片、文字、表格文字、显式几何和图片字节；组合坐标变换、母版继承、图表、SmartArt、动画和形状堆叠顺序未完整比较。'};
export class OfficeEngine {
  constructor(differ){this.differ=differ;this.pdf=new PdfEngine(differ);}
  async extract(bytes,name){const format=name.split('.').pop().toLowerCase();if(format==='pdf')return this.pdf.extract(bytes,name);const pkg=new OfficePackage(bytes),kind={docx:['w','document'],xlsx:['s','workbook'],pptx:['p','presentation']}[format];if(!kind||!is(pkg.xml(pkg.main()),...kind))throw new Error('文件内容与 Office 扩展名不一致，或使用了尚未支持的 OOXML 命名空间。');return format==='docx'?extractWord(pkg,name):format==='xlsx'?extractExcel(pkg,name):extractPowerPoint(pkg,name);}
  async compare(oldBytes,newBytes,oldName,newName,{onProgress}={}){const format=oldName.split('.').pop().toLowerCase();if(!['docx','xlsx','pptx','pdf'].includes(format)||format!==newName.split('.').pop().toLowerCase())throw new Error('请选择同一种格式的 .docx、.xlsx、.pptx 或 .pdf 文件。');
    if(format==='pdf')return this.pdf.compare(oldBytes,newBytes,oldName,newName,{onProgress});
    await onProgress?.('before');const before=await this.extract(oldBytes,oldName);
    await onProgress?.('after');const after=await this.extract(newBytes,newName);
    const [oldHash,newHash]=await Promise.all([sha256(oldBytes),sha256(newBytes)]);
    await onProgress?.('diff');
    const changes=semanticDiff(before,after,this.differ,{compareCachedFormulaValues:format==='xlsx'}),aspects={docx:['existence','text','style'],xlsx:['existence','value','formula','style'],pptx:['existence','text','style','geometry','order','media']}[format];
    // strongKey is a matching hint; it is not part of the public preview view.
    for(const n of [...before.nodes,...after.nodes])delete n.strongKey;
    const changeSet={schemaVersion:'1.0.0',id:'compare-'+oldHash.slice(0,8)+'-'+newHash.slice(0,8),format,documents:{before:{name:oldName,sha256:oldHash},after:{name:newName,sha256:newHash}},comparison:{profile:'office-semantic-v1',aspects,...(format==='xlsx'?{compareCachedFormulaValues:true}:{})},coverage:{status:'partial',issues:[{code:'unsupportedFeature',side:'both',message:limitations[format]},{code:'unsupportedFeature',side:'both',message:'示例匹配采用原生标识或行列标签，并对无标识段落使用内容相似度；跨文件重建、重复标识和复杂重排的身份匹配尚未完整支持。'}]},changes};
    if(format==='xlsx')changeSet.extensions={'ppt2wiki.impact':buildFormulaImpacts({before,after,changeSet})};
    return {changeSet,before,after};
  }
  async annotate(bytes,set,side){return set.format==='pdf'?this.pdf.annotate(bytes,set,side):annotate(bytes,set,side);}
  async disposeResult(result){if(result?.changeSet.format==='pdf')await Promise.all([this.pdf.dispose(result.before),this.pdf.dispose(result.after)]);}
}
