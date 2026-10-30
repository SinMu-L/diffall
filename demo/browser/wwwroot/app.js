import {characterDiff, renderCharacterText} from './text-diff-view.js';
import {formatNumberText} from './value-format.js';
import {CharacterDiffer} from './engine/character-diff.js';
import {OfficeEngine} from './engine/office-engine.js';
import {sheetAxis, columnName, cellPosition, sourceKey, indexReview, revealIn} from './spreadsheet-review.js';
import {indexImpacts,impactLabel,relatedPositions,referencedAreas} from './impact-review.js';
import {pdfPageKey,renderPdfPreview} from './pdf-review.js';
import {changeGroup as groupOf} from './change-types.js';
const $ = id => document.getElementById(id);
const esc = value => String(value ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const icons = {
 layers:'<rect x="4" y="3" width="13" height="15" rx="2"/><path d="M8 7h5M8 11h5M8 15h3M8 21h11a2 2 0 0 0 2-2V8"/>',
 code:'<path d="m8 7-5 5 5 5m8-10 5 5-5 5m-3-13-2 16"/>',
 presentation:'<rect x="3" y="4" width="18" height="13" rx="2"/><path d="M12 17v4m-4 0h8M7 9h10M7 13h6"/>',
 document:'<path d="M14 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9zM14 3v6h6M8 13h8M8 17h5"/>',
 chevron:'<path d="m9 5 7 7-7 7"/>',
 spark:'<path d="m12 3 2.5 6.5L21 12l-6.5 2.5L12 21l-2.5-6.5L3 12l6.5-2.5zM20 3v4m-2-2h4"/>',
 shield:'<path d="m12 3 8 3v6c0 5-8 9-8 9s-8-4-8-9V6zM8.5 12l2.5 2.5 4.5-5"/>',
 download:'<path d="M12 3v12m-4-4 4 4 4-4M4 16v4h16v-4"/>',
 upload:'<path d="M12 15V3m-4 4 4-4 4 4M4 16v4h16v-4"/>',
 swap:'<path d="M4 8h16m-4-4 4 4-4 4M20 16H4m4-4-4 4 4 4"/>',
 compare:'<path d="M7 3v18m10-18v18M3 7l4-4 4 4m2 10 4 4 4-4"/>',
 text:'<path d="M4 5h16M4 10h16M4 15h10M4 20h7m11-7v8m-3-5 3-3 3 3"/>',
 move:'<path d="M12 3v18M3 12h18M8 7l4-4 4 4M8 17l4 4 4-4M7 8l-4 4 4 4m10-8 4 4-4 4"/>',
 plusminus:'<rect x="3" y="4" width="18" height="16" rx="2"/><path d="M7 9h4M9 7v4m5 5h4"/>',
 columns:'<rect x="3" y="4" width="18" height="16" rx="2"/><path d="M12 4v16M6 8h3m6 0h3M6 12h3m6 0h3"/>',
 tree:'<rect x="3" y="3" width="7" height="5" rx="1"/><path d="M6 8v10h8M6 12h8"/><rect x="14" y="10" width="7" height="5" rx="1"/><rect x="14" y="17" width="7" height="4" rx="1"/>',
 info:'<circle cx="12" cy="12" r="9"/><path d="M12 11v6m0-10h.01"/>',
 check:'<path d="m5 12 4 4L19 6"/>',
 sliders:'<path d="M4 7h16M4 17h16M8 4v6m8 4v6"/>',
 link:'<path d="m10 13 4-4m-6 7-2 2a4 4 0 0 1-6-6l5-5a4 4 0 0 1 6 0m2 2 2-2a4 4 0 0 1 6 6l-5 5a4 4 0 0 1-6 0" transform="translate(2 -1)"/>',
 copy:'<rect x="8" y="8" width="12" height="13" rx="2"/><path d="M15 8V3H3v13h5"/>',
 close:'<path d="m6 6 12 12M18 6 6 18"/>',
 arrow:'<path d="M4 12h16m-5-5 5 5-5 5"/>'
};
const icon = name => `<svg class="icon" viewBox="0 0 24 24" aria-hidden="true">${icons[name] || icons.document}</svg>`;
function hydrateIcons(root=document){root.querySelectorAll('[data-icon]').forEach(el=>{el.innerHTML=icon(el.dataset.icon);delete el.dataset.icon;});}
hydrateIcons();

const samples = {
 'pdf-budget':{title:'PDF 预算审阅',format:'pdf',description:'文字、混合样式、位置和页面增删与调序。'},
 'ppt-budget':{title:'预算汇报修订',format:'pptx',description:'同一张幻灯片，追踪金额、表格、位置、增删与页序的变化。'},
 'excel-budget':{title:'部门预算调整',format:'xlsx',description:'插入新部门行后，通过行列标签匹配预算，识别数值与公式修改。'},
 'excel-formulas':{title:'公式与连锁变化',format:'xlsx',description:'修改 B5 后，查看范围合计、连锁计算和跨表结果的变化，并追溯来源。'},
 'word-budget':{title:'预算说明修订',format:'docx',description:'从正文段落中，找到金额修改、加粗和新增的执行说明。'}
};
const formats = {pdf:{label:'PDF',css:'pdf',letter:'PDF',mime:'application/pdf'},pptx:{label:'POWERPOINT',css:'ppt',letter:'P',mime:'application/vnd.openxmlformats-officedocument.presentationml.presentation'},xlsx:{label:'EXCEL',css:'excel',letter:'X',mime:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'},docx:{label:'WORD',css:'word',letter:'W',mime:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'}};
const state = {files:{before:null,after:null},result:null,view:'preview',filter:'all',panel:'changes',impactFilter:'all',rangeEntity:null,selected:null,page:null,showHighlights:true,busy:false,action:null,urls:[]};
const excelReview = () => ({mode:'changes', before:{rows:0,columns:0,expandedRows:[],expandedColumns:[]},after:{rows:0,columns:0,expandedRows:[],expandedColumns:[]}});
let reviewIndex, excel=excelReview(), previewRevision=0;
const aspectLabel = change => change.impact?impactLabel(change.impact):({text:'字符修改',value:reviewIndex?.impact?.byChange.has(change.id)?'计算结果变化':state.result?.changeSet.format==='xlsx'?'输入值变化':change.before?.value.type==='text'&&change.after?.value.type==='text'?'字符修改':'数值修改',formula:'公式修改',style:change.property.endsWith('/bold')?'加粗变化':change.property.endsWith('/italic')?'斜体变化':change.property.endsWith('/fontFamily')?'字体变化':change.property.endsWith('/color')?'颜色变化':'字号变化',geometry:'位置 / 尺寸',order:'页面调序',parent:state.result?.changeSet.format==='pdf'?'跨页移动':'层级调整',media:'图片替换',existence:change.target.kind==='page'?(change.operation==='add'?'新增页面':'删除页面'):change.operation==='add'?'新增元素':'删除元素',metadata:'属性变化'}[change.aspect]||change.aspect);
const short = (value,n=45)=>String(value ?? '').length>n?String(value).slice(0,n)+'…':String(value??'');
const color = value=>/^[0-9a-f]{6}$/i.test(value)?'#'+value:'#243b53';
const num = value=>Number.isFinite(Number(value))?Number(value):0;

function toast(message){$('toast').textContent=message;$('toast').hidden=false;clearTimeout(toast.timer);toast.timer=setTimeout(()=>$('toast').hidden=true,3200);}
function error(message){$('error-message').textContent=message;$('error-message').hidden=!message;}
function busy(value,action='compare'){
 state.busy=value;state.action=value?action:null;document.body.classList.toggle('busy',value);
 const hasFiles=!!state.files.before&&!!state.files.after;
 ['replace-old','replace-new','old-input','new-input','toggle-changes'].forEach(id=>$(id).disabled=value);
 $('swap-files').disabled=value||!hasFiles;
 $('compare').disabled=value||!hasFiles||document.body.dataset.engine!=='ready';
 ['export-json','export-before','export-after'].forEach(id=>$(id).disabled=value||!state.result);
 $('compare').innerHTML=icon('compare')+(value?({compare:'正在对比…',export:'正在导出…',read:'正在读取…'}[action]||'正在处理…'):state.result?'重新对比':'开始对比');
 $('workbench').inert=value;$('workbench').setAttribute('aria-busy',String(value));
}
// A frame followed by a task lets the browser paint before synchronous OOXML work.
const nextPaint=()=>new Promise(resolve=>document.hidden?setTimeout(resolve,0):requestAnimationFrame(()=>setTimeout(resolve,0)));
const comparisonStages=['before','after','diff','preview'];
async function reportProgress(stage,message){
 clearTimeout(finishProgress.timer);
 const index=comparisonStages.indexOf(stage);
 $('processing').hidden=false;$('processing').dataset.status='running';$('processing-bar').removeAttribute('value');
 $('processing-status').textContent=message||({before:'正在解析旧版文件…',after:'正在解析新版文件…',diff:'正在匹配内容并计算差异…',preview:'正在生成对照预览…'}[stage]||'正在准备对比…');
 if(!state.result){$('empty-title').textContent='正在处理文件';$('empty-description').textContent='完成后会自动展示对比结果，请稍候。';}
 $('processing-count').textContent=index<0?'':`第 ${index+1} / ${comparisonStages.length} 步`;
 document.querySelectorAll('[data-stage]').forEach(item=>{const position=comparisonStages.indexOf(item.dataset.stage);item.classList.toggle('active',position===index);item.classList.toggle('complete',index>=0&&position<index);if(position===index)item.setAttribute('aria-current','step');else item.removeAttribute('aria-current');});
 await nextPaint();
}
function finishProgress(message,failed=false){
 $('processing').dataset.status=failed?'error':'complete';$('processing-status').textContent=message;
 $('processing-count').textContent=failed?'可重新尝试':'4 / 4 步完成';$('processing-bar').max=4;$('processing-bar').value=failed?0:4;
 document.querySelectorAll('[data-stage]').forEach(item=>{item.classList.remove('active');item.classList.toggle('complete',!failed);item.removeAttribute('aria-current');});
 if(!failed)finishProgress.timer=setTimeout(()=>$('processing').hidden=true,1400);
}
function clearComparison(){
 const previous=state.result;if(previous)engineReady.then(engine=>engine?.disposeResult(previous)).catch(()=>{});
 state.result=null;state.selected=null;state.page=null;state.rangeEntity=null;state.panel='changes';state.filter='all';state.impactFilter='all';reviewIndex=null;previewRevision++;
 $('workbench').hidden=true;$('comparison-stats').hidden=true;$('comparison-empty').hidden=false;
 const hasFiles=state.files.before&&state.files.after;
 $('empty-title').textContent=hasFiles?'文件已就绪，点击开始对比':'选择两个版本，查看文件差异';
 $('empty-description').textContent=hasFiles?'对比完成后，将在这里展示新旧版本与差异。':'支持 Word、Excel、PowerPoint 和 PDF，文件仅在浏览器内处理。';
 ['stat-total','stat-content','stat-layout','stat-existence'].forEach(id=>$(id).textContent='—');
 $('changes-count').textContent='0';$('change-list').innerHTML='';$('anchor-inspector').innerHTML='';$('json-code').textContent='';$('structure-view').innerHTML='';$('timing').textContent='';$('page-select').innerHTML='';
 $('more-tools').open=false;document.querySelector('.change-details').open=false;
 for(const side of ['before','after']){$(side+'-preview').innerHTML='';$(side+'-preview').classList.remove('spreadsheet-surface','pdf-surface');}
 showView('preview');
}
function blobDownload(bytes,name,type){const url=URL.createObjectURL(new Blob([bytes],{type}));const anchor=document.createElement('a');anchor.href=url;anchor.download=name;anchor.click();setTimeout(()=>URL.revokeObjectURL(url),5000);}
function setFiles(){
 state.urls.forEach(URL.revokeObjectURL);state.urls=[];
 if(!state.result){$('sheet-review-controls').hidden=true;$('impact-summary').hidden=true;$('impact-tabs').hidden=true;$('impact-outcomes').hidden=true;$('calculation-filter').hidden=true;$('preview-navigation').innerHTML='';for(const side of ['before','after'])$(side+'-sheet-tools').hidden=true;}
 for(const [side,prefix] of [['before','old'],['after','new']]){
  const file=state.files[side];if(!file)continue;
  const ext=file.name.split('.').pop().toLowerCase(), f=formats[ext]||formats.pptx;
  $(prefix+'-name').textContent=file.name;$(prefix+'-size').textContent=`${(file.bytes.length/1024).toFixed(1)} KB · ${ext.toUpperCase()}`;
  $(prefix+'-file-icon').className='file-icon '+f.css;$(prefix+'-file-icon').textContent=f.letter;
  $('replace-'+prefix).innerHTML='替换'+icon('upload');$('replace-'+prefix).setAttribute('aria-label',side==='before'?'替换旧版文件':'替换新版文件');$(prefix+'-name').title=file.name;
  const url=URL.createObjectURL(new Blob([file.bytes],{type:f.mime}));state.urls.push(url);$('download-'+prefix).href=url;$('download-'+prefix).download=file.name;
 }
}

const engineReady=(async()=>{
 try{
  const differ=await CharacterDiffer.load();
  const engine=new OfficeEngine(differ);document.body.dataset.engine='ready';
  $('engine-status').innerHTML='<span class="status-dot"></span><span>本地处理</span>';busy(state.busy,state.action);
  return engine;
 }catch(cause){document.body.dataset.engine='failed';$('engine-status').innerHTML='<span class="status-dot failed"></span><span>加载失败</span>';error('无法启动比较引擎，请刷新页面重试。'+cause.message);return null;}
})();

async function loadSample(id){
 if(state.busy)return;busy(true);error('');
 try{
  await reportProgress('preparing','正在读取示例文件…');
  const sample=samples[id];
  const files=await Promise.all(['old','new'].map(async side=>{const name=`${id}-${side}.${sample.format}`;const response=await fetch('samples/'+name);if(!response.ok)throw new Error('示例文件加载失败。');return{name,bytes:new Uint8Array(await response.arrayBuffer())};}));
  state.files.before=files[0];state.files.after=files[1];clearComparison();setFiles();
  await runComparison();
 }catch(cause){clearComparison();finishProgress('示例对比失败',true);error(cause.message);}finally{busy(false);}
}

async function runComparison(){
 clearComparison();await reportProgress('preparing');
 const engine=await engineReady;
 if(!engine)throw new Error('比较引擎加载失败，请刷新页面重试。');
 const before=state.files.before,after=state.files.after;if(!before||!after)throw new Error('请先选择两个文件。');
 if(before.name.split('.').pop().toLowerCase()!==after.name.split('.').pop().toLowerCase())throw new Error('请选择同一种格式的旧版和新版文件。');
 const start=performance.now();
 state.result=await engine.compare(before.bytes,after.bytes,before.name,after.name,{onProgress:reportProgress});
 await reportProgress('preview');state.filter='all';state.panel='changes';state.impactFilter='all';state.rangeEntity=null;state.selected=state.result.changeSet.changes[0]?.id??null;
 reviewIndex=indexReview(state.result);reviewIndex.impact=indexImpacts(state.result);for(const [id,entry] of reviewIndex.impact.byId)reviewIndex.byId.set(id,entry);excel=excelReview();
 const selected=state.result.changeSet.changes[0], locator=selected?.after?.source?.locator||selected?.before?.source?.locator;
 state.page=state.result.changeSet.format==='pdf'?pdfPageKey(state.result,selected)||state.result.after.pages[0]?.key:locator?.slideId?'slide:'+locator.slideId:locator?.sheetId?'sheet:'+locator.sheetId:state.result.after.pages[0]?.key||state.result.before.pages[0]?.key;
 const changes=state.result.changeSet.changes;
 $('stat-total').textContent=changes.length;$('stat-content').textContent=changes.filter(c=>groupOf(c)==='content').length;$('stat-layout').textContent=changes.filter(c=>groupOf(c)==='layout').length;$('stat-existence').textContent=changes.filter(c=>groupOf(c)==='existence').length;
 $('changes-count').textContent=changes.length;
 const isExcel=state.result.changeSet.format==='xlsx',analysis=reviewIndex.impact;
 $('impact-tabs').hidden=!isExcel;$('calculation-filter').hidden=!isExcel;$('impact-count').textContent=analysis.items.length;$('impact-summary').hidden=!isExcel;
 $('impact-limitations').open=false;$('impact-limitations-count').textContent=analysis.dependencyLimitations.count;
 $('impact-limitations-list').innerHTML=analysis.dependencyLimitations.details.map(detail=>`<li>${esc(detail)}</li>`).join('');
 $('impact-summary').innerHTML=analysis.unsupported?'影响分析版本暂不支持，当前仅展示实际变化。':`<span class="fx-symbol">ƒx</span><span><strong>${analysis.stats.changed}</strong> 个单元格的计算结果变化，已计入变化总数</span><button data-panel="impacts" data-impact-outcome="unknown">${analysis.stats.unknown} 个结果待确认</button>${analysis.stats.unchanged?'<span>'+analysis.stats.unchanged+' 个结果未变</span>':''}${analysis.stats.derived?'<span>'+analysis.stats.derived+' 个重算分析</span>':''}<small>文件保存结果 · 原因尚未验证</small>`;
 $('format-label').innerHTML='<span class="mini-dot"></span>'+formats[state.result.changeSet.format].label;
 const isPdf=state.result.changeSet.format==='pdf';$('preview-kind').textContent=isPdf?'PDF 原始页面 · 区域高亮':'结构化预览';$('export-description').textContent=isPdf?'下载含差异区域标记的 PDF 副本':'下载含差异标记的 Office 文件';$('json-filename').textContent=isPdf?'pdf-change.v1.json':'office-change.v1.json';
 $('page-select-label').textContent=isExcel?'工作表':'页面';$('page-select').setAttribute('aria-label',isExcel?'选择工作表':'选择页面');
 $('coverage-text').textContent=[...state.result.changeSet.coverage.issues.map(issue=>issue.message),...(analysis.dependencyLimitations.count?[`${analysis.dependencyLimitations.count} 个公式的引用未完整解析，具体原因见计算影响中的“分析限制”；仅有解析限制的单元格不计入计算影响。`]:[])].join(' ');
 const pages=[...state.result.after.pages,...state.result.before.pages].filter((page,index,all)=>all.findIndex(p=>p.key===page.key)===index);
 $('page-select').innerHTML=pages.map(page=>`<option value="${esc(page.key)}">${esc(short(page.name,20))}</option>`).join('');$('page-select').value=state.page;
 document.querySelectorAll('[data-filter]').forEach(button=>button.classList.toggle('active',button.dataset.filter==='all'));
 document.body.dataset.totalChanges=changes.length;
 $('workbench').hidden=false;$('comparison-stats').hidden=false;$('comparison-empty').hidden=true;
 renderChanges();await renderPreviews({locate:true,preserve:false});
 await nextPaint();const elapsed=performance.now()-start;$('timing').textContent=elapsed<1000?`${Math.round(elapsed)} ms`:`${(elapsed/1000).toFixed(1)} 秒`;
 finishProgress(changes.length?`对比完成，发现 ${changes.length} 处变化`:'对比完成，已比较范围内未发现变化');
}

function humanValue(value,change){
 if(!value)return'不存在';const data=value.data;
 if(value.type==='null')return'空值';if(value.type==='boolean')return change?.property.endsWith('/bold')?(data?'加粗':'常规'):(data?'是':'否');
 if(value.type==='number')return formatNumberText(data);
 if(value.type==='asset')return '图片 '+data.sha256.slice(0,8);
 if(value.type==='array'&&change?.aspect==='style'){const display=v=>v==null?'未知':typeof v==='boolean'?(v?'是':'否'):String(v);return [...new Set(data.filter(r=>r.value!=null).map(r=>display(r.value)))].join(' / ')||'未知';}
 if(value.type==='object'){
  if(change?.aspect==='geometry')return data.unit==='pt'?`${data.x}, ${data.y} · ${data.width} × ${data.height} pt · ${data.rotation}°`:`x ${Math.round(data.x/9525)} · y ${Math.round(data.y/9525)}`;
  if(change?.operation==='move')return'第 '+(data.index+1)+' 页';
  return short(data.text||data.scalar?.data||({shape:'形状',cell:'单元格',paragraph:'段落',slide:'幻灯片',page:'PDF 页面',textRange:'文字区域',table:'表格'}[data.kind]||'对象'),40);
 }
 return data===''?'空字符串':short(data,60);
}
function anchorLabel(source,side){if(!source)return'不存在';const l=source.locator;if(l.type==='pdf')return`第 ${l.pageIndex+1} 页 · 区域 ${l.rect.map(n=>Math.round(n)).join(', ')}`;if(l.address){const name=(side?state.result?.[side]?.pages:[...(state.result?.after.pages||[]),...(state.result?.before.pages||[])])?.find(page=>page.key==='sheet:'+l.sheetId)?.name;return`${name||'Sheet '+l.sheetId} · ${l.address}`;}if(l.slideId)return`Slide ${l.slideId}`+(l.shapePath?' · Shape '+l.shapePath.join('/'):'')+(l.cell?` · Cell ${l.cell.row},${l.cell.column}`:'');return'XML '+(l.elementPath?.join('/')||'root');}
function filteredChanges(){if(state.panel==='impacts')return(reviewIndex?.impact?.items||[]).filter(c=>state.impactFilter==='all'||c.impact.outcome===state.impactFilter);return(state.result?.changeSet.changes||[]).filter(change=>state.filter==='all'||(state.filter==='calculation'?reviewIndex?.impact?.byChange.has(change.id):groupOf(change)===state.filter));}
function changeValues(change){
 if(change.impact){const entity=reviewIndex.impact.entities.get(change.target.entityId);return`<span class="change-values">${['before','after'].map(side=>{const o=entity[side]?.observations.find(o=>o.property===change.property);return`<span class="${change.impact.outcome==='changed'?(side==='before'?'old-value':'new-value'):'observed-value'}">${o?.status==='known'?esc(humanValue(o.value)): '未取得结果'}</span>`;}).join(icon('arrow'))}</span>`;}
 const diff=characterDiff(change);
 if(diff)return`<span class="change-values character-values">${['before','after'].map(side=>`<span class="inline-side"><span class="diff-side-label">${side==='before'?'旧':'新'}</span><span class="inline-value">${renderCharacterText(humanValue(change[side].value,change),diff,side,{rawText:change[side].value.data})}</span></span>`).join('')}</span>`;
 return`<span class="change-values">${change.before?`<span class="old-value">${esc(humanValue(change.before.value,change))}</span>`:''}${change.before&&change.after?icon('arrow'):''}${change.after?`<span class="new-value">${esc(humanValue(change.after.value,change))}</span>`:''}</span>`;
}
function renderChanges(){
 if(!state.result)return;
 const scrollTop=$('change-list').scrollTop;
 const changes=filteredChanges();
 const renderItem=change=>{const impact=change.impact||reviewIndex.impact.byChange.get(change.id);return`<button class="change-item ${state.selected===change.id?'selected':''}" data-change="${esc(change.id)}" aria-pressed="${state.selected===change.id}"><span class="change-item-top"><span class="change-num">${change.impact?'ƒx':esc(change.id.replace('change-','#'))}</span><span class="change-tag ${impact?'calculation':groupOf(change)} ${impact?.outcome==='unknown'?'uncertain':''}">${esc(aspectLabel(change))}</span>${characterDiff(change)?`<span class="character-count">${characterDiff(change).edits.length} 处</span>`:''}</span><strong class="change-name" title="${esc(change.target.label)}">${esc(change.target.label)}</strong>${changeValues(change)}${impact?'<span class="impact-caveat">'+(impact.evidence.level==='verified'?'已提供验证记录':'原因待确认')+'</span>':''}<span class="change-item-foot">${esc(anchorLabel(change.after?.source||change.before?.source))}</span></button>`;};
 const visible=changes.filter(c=>c.impact?.outcome!=='unchanged'),unchanged=changes.filter(c=>c.impact?.outcome==='unchanged');
 $('change-list').innerHTML=changes.length?visible.map(renderItem).join('')+(unchanged.length?`<details class="unchanged-impacts" ${unchanged.some(c=>c.id===state.selected)?'open':''}><summary>结果未变 · ${unchanged.length} 个对象</summary>${unchanged.map(renderItem).join('')}</details>`:''):`<div class="empty-changes">${state.panel==='impacts'?'本次没有可展示的计算影响。':state.result.changeSet.changes.length?'当前分类没有变化。':'已处理范围没有识别到变化。'}<br>完整覆盖情况见下方说明。</div>`;
 const change=reviewIndex.byId.get(state.selected);
 $('change-list').scrollTop=scrollTop;
 $('anchor-inspector').innerHTML=(change?renderImpactInspector(change):'')+`<div>${icon('link')}原文件定位</div>`+(change?['before','after'].map(side=>`<div class="anchor-line"><span>${side==='before'?'旧版':'新版'}</span><code>${esc(anchorLabel(change[side]?.source,side))}</code></div>`).join(''):'<p>选择一处变化，查看其定位锚点。</p>');
 document.querySelectorAll('[data-panel]').forEach(button=>button.setAttribute('aria-selected',String(button.dataset.panel===state.panel)));document.querySelector('.change-filters').hidden=state.panel==='impacts';
 $('impact-outcomes').hidden=state.panel!=='impacts';$('impact-outcomes').innerHTML=[['all','全部影响'],['changed','结果变化'],['unknown','待确认'],['unchanged','结果未变']].map(([key,label])=>`<button data-panel="impacts" data-impact-outcome="${key}" aria-pressed="${state.impactFilter===key}">${label}</button>`).join('');
 $('impact-limitations').hidden=state.panel!=='impacts'||!reviewIndex.impact.dependencyLimitations.count;
}

function renderImpactInspector(entry){
 const index=reviewIndex.impact,entityId=entry.target.entityId,entity=index.entities.get(entityId),impact=entry.impact||index.byEntity.get(entityId),downstream=index.relatedDownstream(entityId);
 if(!entity||!impact&&!downstream.length)return'';
 const entityButton=(id,label)=>`<button class="impact-link" data-inspect-entity="${esc(id)}">${esc(label)}</button>`;
 const location=id=>{const e=index.entities.get(id);return anchorLabel(e?.after?.source||e?.before?.source);};
 let html='<section class="impact-inspector">';
 if(impact){
  const computed=Object.values(entity).some(s=>s?.observations?.some(o=>['recalculated','rendered'].includes(o.origin)));
  html+=`<h3><span class="fx-symbol">ƒx</span>${esc(impactLabel(impact))}</h3><small class="observation-label">${computed?'重算分析值':'文件保存值'}</small><div class="impact-values">${['before','after'].map(side=>{const o=entity[side]?.observations.find(o=>o.property===impact.property);return`<span>${o?.status==='known'?esc(humanValue(o.value)):'未取得结果'}</span>`;}).join('<span class="value-arrow">→</span>')}</div>`;
  if(computed)html+='<p class="impact-note">以上为重算分析值，表格保留原文件中的值。</p>';
  const old=entity.before?.formula,next=entity.after?.formula;
  if(old!=null||next!=null)html+=old===next?`<div class="formula-context"><span>公式未变</span><code>=${esc(next||old)}</code></div>`:`<div class="formula-context"><span>公式变化</span>${[old,next].map(f=>`<code>${f==null?'无公式':'='+esc(f)}</code>`).join(' → ')}</div>`;
  html+=`<p class="impact-evidence">${impact.evidence.level==='verified'?'已提供因果验证记录':'存在引用或结果差异，原因待确认'}</p>`;
  html+=`<h4>变化来源</h4><div class="impact-causes">${impact.causes.map(c=>{if(c.kind==='unresolved')return`<p>${esc(c.detail)}</p>`;if(c.kind==='environment')return`<p>计算环境：${esc(c.name)}</p>`;if(c.kind==='dependency')return'<p>引用范围或目标发生变化。</p>';const change=state.result.changeSet.changes.find(v=>v.id===c.changeId);if(!change)return'<p>来源记录不可用。</p>';const path=c.dependencyIds.map(id=>index.dependencies.get(id)).filter(Boolean),refs=path.map(d=>d.after?.expression||d.before?.expression).filter(Boolean);return`<div class="impact-source">${change.target.entityId===entityId?'<strong>本单元格公式变化</strong>':entityButton(change.target.entityId,location(change.target.entityId))}${changeValues(change)}${refs.length?`<small>引用路径：${esc(refs.join(' → '))}</small>`:''}</div>`;}).join('')}</div>`;
  if([...index.dependencies.values()].some(d=>d.dependentEntityId===entityId))html+=`<button class="range-button" data-show-range="${esc(entityId)}" aria-pressed="${state.rangeEntity===entityId}">${state.rangeEntity===entityId?'收起引用范围':'查看引用范围'}</button>`;
  html+=`<p class="impact-note">${computed?'重算结果单独记录。':impact.evidence.level==='verified'?'已提供计算结果复核记录。':impact.outcome==='unknown'?'缓存相同或缺失，尚不能确认重算结果。':'文件保存的计算结果，缓存尚未重算验证。'}</p>`;
 }
 if(downstream.length)html+=`<h4>关联 ${downstream.filter(i=>i.outcome==='changed').length} 处结果变化${downstream.some(i=>i.outcome==='unknown')?' · 含待确认影响':''}</h4><div class="impact-downstream">${downstream.slice(0,8).map(i=>entityButton(i.entityId,location(i.entityId)+' · '+impactLabel(i))).join('')}${downstream.length>8?`<details><summary>查看其余 ${downstream.length-8} 项</summary>${downstream.slice(8).map(i=>entityButton(i.entityId,location(i.entityId)+' · '+impactLabel(i))).join('')}</details>`:''}</div>`;
 return html+'</section>';
}

function nodeMarker(node,side){
 const changes=reviewIndex.changes[side].get(sourceKey(node.source))||[];
 const selected=changes.find(change=>change.id===state.selected),first=selected||changes[0];
 const inlineChange=changes.find(change=>characterDiff(change)&&change.aspect!=='formula');
 const objectChange=selected||changes.find(change=>!characterDiff(change))||first;
 const impact=reviewIndex.impact.atSource[side].get(sourceKey(node.source)),selectedEntry=reviewIndex.byId.get(state.selected),entitySelected=selectedEntry?.target.entityId===node.entityId,analysisId=impact?'analysis:'+impact.id:null;
 return {hit:!!first&&(state.showHighlights||!!selected),selected:!!selected||entitySelected,id:entitySelected?state.selected:first?.id||analysisId,group:objectChange?groupOf(objectChange):null,impact,inlineChange,formulaChange:changes.find(change=>change.aspect==='formula'),paintObject:!!first&&(state.showHighlights||!!selected)&&(!inlineChange||!!objectChange&&!characterDiff(objectChange))};
}
const markerGroup = mark => mark.group?`data-change-group="${mark.group}"`:'';
function nodeText(text,mark,side){
 if(mark.inlineChange&&state.showHighlights)return renderCharacterText(text,characterDiff(mark.inlineChange),side,{rawText:mark.inlineChange[side].value.data});
 if(mark.hit&&state.showHighlights){const change=state.result.changeSet.changes.find(change=>change.id===mark.id);if(change?.operation==='add'||change?.operation==='remove'){const tag=side==='before'?'del':'ins';return`<${tag} class="char-diff ${side}">${esc(text)}</${tag}>`;}}
 return esc(text);
}
function renderPpt(view,side){
 const page=view.pages.find(page=>page.key===state.page);if(!page)return'<div class="preview-placeholder"><strong>此版本没有该页面</strong></div>';
 const scale=1280/page.width;const nodes=view.nodes.filter(node=>node.pageKey===page.key&&node.geometry);
 const slide=view.nodes.find(node=>node.key===page.key);const pageMark=slide?nodeMarker(slide,side):{};
 const html=nodes.map(node=>{
  const g=node.geometry;const x=g.x*scale,y=g.y*scale,width=g.width*scale,height=g.height*scale;
  const mark=nodeMarker(node,side),data=(mark.id?`data-change="${esc(mark.id)}" class="node-click" `:'')+markerGroup(mark);
  const background=node.fill!=='transparent'?`<rect x="${x}" y="${y}" width="${width}" height="${height}" fill="${color(node.fill)}"/>`:'';
  let content='';
  if(node.table){const shapePath=JSON.stringify(node.source.locator.shapePath);content=`<foreignObject x="${x}" y="${y}" width="${width}" height="${height}"><table xmlns="http://www.w3.org/1999/xhtml" class="slide-table">${node.table.map((row,r)=>`<tr>${row.map((text,c)=>{const cell=view.nodes.find(cell=>cell.pageKey===page.key&&cell.kind==='cell'&&JSON.stringify(cell.source.locator.shapePath)===shapePath&&cell.source.locator.cell?.row===r&&cell.source.locator.cell?.column===c);const m=cell?nodeMarker(cell,side):{};return`<td class="${m.paintObject?'node-hit '+side:''} ${m.id?'node-click':''}" ${markerGroup(m)} ${m.id?`data-change="${esc(m.id)}"`:''}>${nodeText(text,m,side)}</td>`;}).join('')}</tr>`).join('')}</table></foreignObject>`;}
  else if(node.imageData)content=`<image x="${x}" y="${y}" width="${width}" height="${height}" href="${esc(node.imageData)}"/>`;
  else if(node.text!==null)content=`<foreignObject x="${x}" y="${y}" width="${width}" height="${height}"><div xmlns="http://www.w3.org/1999/xhtml" class="shape-text" style="font-size:${num(node.fontSize)*4/3}px;color:${color(node.color)};font-weight:${node.bold?650:400}">${nodeText(node.text,mark,side)}</div></foreignObject>`;
  const transform=`translate(${x+width/2} ${y+height/2}) rotate(${num(g.rotation)/60000}) scale(${g.flipH?-1:1} ${g.flipV?-1:1}) translate(${-x-width/2} ${-y-height/2})`;
  return`<g ${data} transform="${transform}">${background}${mark.paintObject?`<rect class="hit-zone ${mark.selected?'node-selected':''}" x="${x-5}" y="${y-5}" width="${width+10}" height="${height+10}" rx="5" stroke-width="3" stroke-dasharray="8 5"/>`:''}${content}</g>`;
 }).join('');
 return`<svg class="slide-canvas ${pageMark.hit?'changed-page':''} ${pageMark.selected?'selected-page':''}" ${markerGroup(pageMark)} ${pageMark.id?`data-change="${esc(pageMark.id)}"`:''} viewBox="0 0 1280 ${page.height*scale}" role="img" aria-label="${esc(page.name)}"><rect width="1280" height="${page.height*scale}" fill="white"/>${html}</svg>`;
}
function renderWord(view,side){return`<article class="word-sheet">${view.nodes.map(node=>{const mark=nodeMarker(node,side);return`<p class="${mark.paintObject?'node-hit '+side:''} ${mark.selected?'selected':''} ${mark.id?'node-click':''}" ${markerGroup(mark)} style="font-size:${Math.min(24,node.fontSize*1.05)}px;font-weight:${node.bold?650:400};color:${color(node.color)}" ${mark.id?`data-change="${esc(mark.id)}"`:''}>${nodeText(node.text,mark,side)||'&nbsp;'}</p>`;}).join('')}</article>`;}
function sheetChanges(pageKey,changes=filteredChanges()){
 return changes.filter(change=>['before','after'].some(side=>'sheet:'+change[side]?.source?.locator?.sheetId===pageKey));
}
function axisPager(axis,side,name){
 return `<div class="sheet-pager"><span>${name==='rows'?'行':'列'} ${axis.start}–${axis.end} / ${axis.total}</span><button data-sheet-page="${name}" data-side="${side}" data-page="${axis.page-1}" aria-label="上一页${name==='rows'?'行':'列'}" ${axis.page===0?'disabled':''}>‹</button><span>${axis.page+1}/${axis.pages}</span><button data-sheet-page="${name}" data-side="${side}" data-page="${axis.page+1}" aria-label="下一页${name==='rows'?'行':'列'}" ${axis.page===axis.pages-1?'disabled':''}>›</button></div>`;
}
function expandButton(gap,side,axis){
 const count=gap.end-gap.start+1,label=axis==='rows'?'行':'列';
 return `<button class="sheet-expand" data-sheet-expand="${axis}" data-side="${side}" data-start="${gap.start}" data-end="${gap.end}" title="展开${Math.min(count,axis==='rows'?50:12)}${label}">${axis==='rows'?`展开 ${gap.start}–${gap.end} 行`:`+${count}列`}</button>`;
}
function renderExcel(view,side,locate){
 const page=view.pages.find(page=>page.key===state.page),tools=$(side+'-sheet-tools');
 if(!page){tools.innerHTML='';return'<div class="preview-placeholder"><strong>此版本没有该工作表</strong></div>';}
 const sheet=reviewIndex.sheets[side].get(page.key)||{cells:new Map(),maxRow:1,maxColumn:1};
 const changes=sheetChanges(page.key),positions=changes.map(change=>cellPosition(change[side]?.source)).filter(Boolean);
 const selected=reviewIndex.byId.get(state.selected),source=selected?.[side]?.source;
 if(selected)positions.push(...relatedPositions(reviewIndex.impact,selected.target.entityId,side,page.key));
 const areas=state.rangeEntity?referencedAreas(reviewIndex.impact,state.rangeEntity,side).filter(area=>reviewIndex.impact.entities.get(area.sheetEntityId)?.[side]?.source.locator.sheetId===page.source.locator.sheetId):[];
 const rangeStart=areas[0]?cellPosition({locator:{address:areas[0].start}}):null;
 if(rangeStart)positions.push({row:Math.min(rangeStart.row,sheet.maxRow),column:Math.min(rangeStart.column,sheet.maxColumn)});
 const target='sheet:'+source?.locator?.sheetId===page.key?cellPosition(source):null;
 const missing=selected&&sheetChanges(page.key,[selected]).length&&!source;
 const nearby=missing?cellPosition(selected[side==='before'?'after':'before']?.source):null;
 const context=target||(nearby?{row:Math.min(nearby.row,sheet.maxRow),column:Math.min(nearby.column,sheet.maxColumn)}:null)||(rangeStart?{row:Math.min(rangeStart.row,sheet.maxRow),column:Math.min(rangeStart.column,sheet.maxColumn)}:null);
 if(target)positions.push(target);
 if(missing&&context)positions.push(context);
 const settings=excel[side];
 const rows=sheetAxis({max:sheet.maxRow,changed:positions.map(p=>p.row),expanded:settings.expandedRows,mode:excel.mode,page:settings.rows,target:locate?context?.row:undefined});
 const columns=sheetAxis({max:sheet.maxColumn,changed:positions.map(p=>p.column),expanded:settings.expandedColumns,mode:excel.mode,page:settings.columns,pageSize:24,target:locate?context?.column:undefined});
 settings.rows=rows.page;settings.columns=columns.page;
 tools.innerHTML=`<div class="sheet-location ${missing?'missing-location':''}"><strong>${esc(page.name)}</strong><span>${target?esc(source.locator.address)+(selected.before&&selected.after?'':' · '+(side==='before'?'已删除':'新增')):missing?'此版本无对应单元格 · '+(side==='before'?'新增':'已删除'):'选择明细定位单元格'}</span></div><div class="sheet-pagination">${axisPager(rows,side,'rows')}${axisPager(columns,side,'columns')}</div><div class="sheet-context-note">${excel.mode==='changes'?`保留相邻数据 · 折叠 ${rows.hidden} 行、${columns.hidden} 列`:`完整数据 · 共 ${sheet.maxRow} 行、${sheet.maxColumn} 列`}</div>`;
 if(areas.length)tools.innerHTML+=`<div class="reference-range-note">引用范围 ${areas.map(a=>esc(a.start+(a.end===a.start?'':':'+a.end))).join('、')} · 当前可见部分 <button data-clear-range>清除</button></div>`;
 let html=`<table class="excel-table" aria-label="${esc(page.name)} ${side==='before'?'原始版本':'修订版本'}"><colgroup><col class="sheet-row-number">${columns.entries.map(c=>`<col class="${typeof c==='number'?'sheet-data-column':'sheet-gap-column'}">`).join('')}</colgroup><thead><tr><th aria-label="行号"></th>${columns.entries.map(c=>`<th scope="col">${typeof c==='number'?columnName(c):expandButton(c,side,'columns')}</th>`).join('')}</tr></thead><tbody>`;
 for(const r of rows.entries){
  if(typeof r!=='number'){html+=`<tr class="sheet-gap-row"><th>⋯</th><td colspan="${columns.entries.length}"><span class="sheet-gap-content">${expandButton(r,side,'rows')}<span>未变化，点击分段展开</span></span></td></tr>`;continue;}
  html+=`<tr><th scope="row">${r}</th>`;
  for(const c of columns.entries){
   if(typeof c!=='number'){html+='<td class="sheet-gap-cell">⋯</td>';continue;}
   const node=sheet.cells.get(`${r}:${c}`),mark=node?nodeMarker(node,side):{};
   let display=node?.display??'';
   if(node?.scalar?.type==='number'||node?.cachedValue?.type==='number')display=formatNumberText(display);
   const formula=node?.formula&&(mark.formulaChange||mark.selected&&mark.impact)?`<span class="sheet-formula">=${mark.formulaChange?renderCharacterText(node.formula,characterDiff(mark.formulaChange),side,{enabled:state.showHighlights}):esc(node.formula)}</span>`:'';
   const inside=areas.some(a=>{const start=cellPosition({locator:{address:a.start}}),end=cellPosition({locator:{address:a.end}});return r>=start.row&&r<=end.row&&c>=start.column&&c<=end.column;});
   const showImpact=mark.impact&&(state.showHighlights||mark.selected)&&(mark.impact.outcome!=='unchanged'||state.panel==='impacts'||mark.selected);
   const badge=showImpact?`<span class="cell-impact-badge ${mark.impact.outcome==='unknown'?'pending':''}" title="${esc(impactLabel(mark.impact))}" aria-label="${esc(impactLabel(mark.impact))}">${mark.impact.outcome==='unknown'?'?':'ƒx'}</span>`:'';
   html+=`<td data-cell="${columnName(c)+r}" ${markerGroup(mark)} ${context?.row===r&&context?.column===c?'data-context-anchor="true"':''} class="${r===1?'header-cell':''} ${mark.hit?'changed-cell':''} ${mark.paintObject?'node-hit '+side:''} ${showImpact?'has-impact impact-'+mark.impact.outcome:''} ${inside?'reference-area':''} ${mark.selected?'selected':''} ${mark.id?'node-click':''}" ${mark.id?`data-change="${esc(mark.id)}"`:''} style="font-weight:${node?.bold?650:400};color:${color(node?.color)}" title="${esc(node?.formula?'='+node.formula:node?.display??'')}">${nodeText(display,mark,side)}${badge}${formula}</td>`;
  }
  html+='</tr>';
 }
 return html+'</tbody></table>';
}
function renderReviewControls(){
 const changes=filteredChanges(),selected=reviewIndex.byId.get(state.selected),index=changes.findIndex(change=>change.id===state.selected),isExcel=state.result.changeSet.format==='xlsx';
 const pages=[...state.result.after.pages,...state.result.before.pages].filter((page,index,all)=>all.findIndex(p=>p.key===page.key)===index);
 $('sheet-review-controls').hidden=!isExcel;
 if(isExcel){
  const changedSheets=pages.filter(page=>sheetChanges(page.key,state.result.changeSet.changes).length).length;
  $('sheet-review-controls').innerHTML=`<div class="sheet-review-heading"><span><strong>${state.result.changeSet.changes.length}</strong> 处变更 · ${changedSheets} 张工作表</span><div class="sheet-mode" role="group" aria-label="表格展示范围"><button data-sheet-mode="changes" aria-pressed="${excel.mode==='changes'}">变更聚焦</button><button data-sheet-mode="all" aria-pressed="${excel.mode==='all'}">完整数据</button></div></div><div class="sheet-chips" aria-label="工作表变更概览">${pages.map(page=>`<button data-sheet="${esc(page.key)}" aria-pressed="${page.key===state.page}">${esc(page.name)}<span>${sheetChanges(page.key).length}</span></button>`).join('')}</div>`;
 }
 $('preview-navigation').innerHTML=`<div class="change-navigation"><button data-change-step="-1" ${index<=0?'disabled':''} aria-label="上一处变更">‹</button><span>${index<0?'0':index+1} / ${changes.length} ${state.panel==='impacts'?'项计算影响':'处变更'}</span><button data-change-step="1" ${!changes.length||index>=changes.length-1?'disabled':''} aria-label="下一处变更">›</button></div><span class="current-change-location">${selected?esc(['before','after'].map(side=>(side==='before'?'旧版：':'新版：')+anchorLabel(selected[side]?.source,side)).join(' → ')):'选择右侧变更明细，自动定位原文'}</span>`;
}
async function renderPreviews({locate=false,preserve=true}={}){
 if(!state.result)return;
 const revision=++previewRevision,isExcel=state.result.changeSet.format==='xlsx',isPdf=state.result.changeSet.format==='pdf',renders=[],positions=[];
 renderReviewControls();
 for(const side of ['before','after']){
  const view=state.result[side],surface=$(side+'-preview'),top=surface.scrollTop,left=surface.scrollLeft;
  positions.push({surface,top,left});
  surface.classList.toggle('spreadsheet-surface',isExcel);surface.classList.toggle('pdf-surface',isPdf);$(side+'-sheet-tools').hidden=!isExcel;
  if(isPdf){const result=state.result;renders.push(engineReady.then(engine=>renderPdfPreview(surface,{view,pageKey:state.page,side,changes:result.changeSet.changes,selected:state.selected,showHighlights:state.showHighlights,engine:engine.pdf})).catch(cause=>{if(revision===previewRevision)error('PDF 页面渲染失败：'+cause.message);}));}
  else surface.innerHTML=state.result.changeSet.format==='pptx'?renderPpt(view,side):state.result.changeSet.format==='docx'?renderWord(view,side):renderExcel(view,side,locate);
  surface.scrollTop=preserve?top:0;surface.scrollLeft=preserve?left:0;
  const pageIndex=view.pages.findIndex(page=>page.key===state.page);$((side==='before'?'old':'new')+'-page-label').textContent=pageIndex>=0?`${String(pageIndex+1).padStart(2,'0')} / ${String(view.pages.length).padStart(2,'0')}`:'—';
 }
 await Promise.all(renders);
 if(revision===previewRevision&&isPdf)for(const {surface,top,left}of positions){surface.scrollTop=preserve?top:0;surface.scrollLeft=preserve?left:0;}
 if(locate)requestAnimationFrame(()=>{
  if(revision!==previewRevision)return;
  const behavior=matchMedia('(prefers-reduced-motion: reduce)').matches?'auto':'smooth';
  for(const side of ['before','after']){const surface=$(side+'-preview');const target=surface.querySelector(`[data-change="${CSS.escape(state.selected||'')}"]`)||surface.querySelector('[data-context-anchor]');revealIn(surface,target,behavior);}
 });
}
function renderStructure(){
 if(!state.result||state.view!=='structure')return;
 $('structure-view').innerHTML=['before','after'].map(side=>`<div class="structure-column"><h3>${side==='before'?'原始版本':'修订版本'} · 提取结果</h3>${state.result[side].nodes.filter(node=>node.pageKey===state.page).map(node=>`<details class="node-detail"><summary><code>${esc(node.kind)}</code>${esc(short(node.text||node.label||node.key,45))}</summary><pre>${esc(JSON.stringify({key:node.key,text:node.text,value:node.scalar,formula:node.formula,geometry:node.geometry,source:node.source},null,2))}</pre></details>`).join('')}</div>`).join('');
}
function selectChange(id){
 const change=reviewIndex?.byId.get(id);if(!change)return;state.selected=id;
 state.rangeEntity=null;
 if(id.startsWith('analysis:')){state.panel='impacts';if(state.impactFilter!=='all'&&change.impact.outcome!==state.impactFilter)state.impactFilter='all';}
 else if(change.operation!=='inspect'&&!filteredChanges().some(c=>c.id===id)){state.panel='changes';state.filter='all';document.querySelectorAll('[data-filter]').forEach(b=>b.classList.toggle('active',b.dataset.filter==='all'));}
 const locator=change.after?.source?.locator||change.before?.source?.locator;
 const page=state.result.changeSet.format==='pdf'?pdfPageKey(state.result,change)||state.page:locator?.slideId?'slide:'+locator.slideId:locator?.sheetId?'sheet:'+locator.sheetId:state.page;
 if(page!==state.page){const mode=excel.mode;excel=excelReview();excel.mode=mode;}
 state.page=page;$('page-select').value=state.page;
 showView('preview');renderChanges();renderPreviews({locate:true});renderStructure();
}
function showView(view){state.view=view;document.querySelectorAll('[data-view]').forEach(tab=>{tab.classList.toggle('active',tab.dataset.view===view);tab.setAttribute('aria-pressed',String(tab.dataset.view===view));});['preview','structure','json'].forEach(name=>$(name+'-view').hidden=name!==view);$('active-view-label').hidden=view==='preview';$('active-view-label').textContent=view==='structure'?'结构视图':view==='json'?'Change JSON':'';if(view==='structure')renderStructure();if(view==='json'&&state.result)$('json-code').textContent=JSON.stringify(state.result.changeSet,null,2);}
function selectPage(page){
 state.page=page;$('page-select').value=page;const mode=excel.mode;excel=excelReview();excel.mode=mode;
 const change=sheetChanges(page)[0];
 if(change){selectChange(change.id);return;}
 state.selected=null;renderChanges();renderPreviews({preserve:false});renderStructure();
}

document.querySelectorAll('[data-view]').forEach(button=>button.addEventListener('click',()=>{showView(button.dataset.view);$('more-tools').open=false;if(state.view==='preview')renderPreviews({locate:true});}));
document.querySelectorAll('[data-filter]').forEach(button=>button.addEventListener('click',()=>{if(!state.result)return;state.filter=button.dataset.filter;document.querySelectorAll('[data-filter]').forEach(other=>other.classList.toggle('active',other===button));const mode=excel.mode;excel=excelReview();excel.mode=mode;const selected=filteredChanges()[0];if(selected)selectChange(selected.id);else{state.selected=null;renderChanges();renderPreviews({preserve:false});}}));
document.addEventListener('click',event=>{const element=event.target.closest('[data-change]');if(element)selectChange(element.dataset.change);});
document.addEventListener('click',event=>{
 if(!state.result)return;
 const panel=event.target.closest('[data-panel]');if(panel){setChangesVisible(true);$('more-tools').open=false;state.panel=panel.dataset.panel;state.rangeEntity=null;state.impactFilter=panel.dataset.impactOutcome||'all';const first=filteredChanges()[0];if(first)selectChange(first.id);else{state.selected=null;renderChanges();renderPreviews();}return;}
 const inspect=event.target.closest('[data-inspect-entity]');if(inspect){const id=reviewIndex.impact.targetId(inspect.dataset.inspectEntity);if(id){state.panel=id.startsWith('analysis:')?'impacts':'changes';state.filter='all';document.querySelectorAll('[data-filter]').forEach(b=>b.classList.toggle('active',b.dataset.filter==='all'));selectChange(id);}return;}
 const range=event.target.closest('[data-show-range]');if(range){state.rangeEntity=state.rangeEntity===range.dataset.showRange?null:range.dataset.showRange;if(state.rangeEntity){const area=referencedAreas(reviewIndex.impact,state.rangeEntity,'after')[0]||referencedAreas(reviewIndex.impact,state.rangeEntity,'before')[0];const e=area&&reviewIndex.impact.entities.get(area.sheetEntityId),source=e?.after?.source||e?.before?.source;if(source){state.page='sheet:'+source.locator.sheetId;$('page-select').value=state.page;}}renderChanges();renderPreviews({locate:true});return;}
 if(event.target.closest('[data-clear-range]')){state.rangeEntity=null;renderChanges();renderPreviews();return;}
 const step=event.target.closest('[data-change-step]');
 if(step){const changes=filteredChanges(),index=changes.findIndex(c=>c.id===state.selected),change=changes[index+Number(step.dataset.changeStep)];if(change){selectChange(change.id);revealIn($('change-list'),$('change-list').querySelector('.selected'));}return;}
 const sheet=event.target.closest('[data-sheet]');if(sheet){selectPage(sheet.dataset.sheet);return;}
 const mode=event.target.closest('[data-sheet-mode]');if(mode){excel.mode=mode.dataset.sheetMode;renderPreviews({locate:true,preserve:false});return;}
 const page=event.target.closest('[data-sheet-page]');if(page){excel[page.dataset.side][page.dataset.sheetPage]=Number(page.dataset.page);renderPreviews();const surface=$(page.dataset.side+'-preview');if(page.dataset.sheetPage==='rows')surface.scrollTop=0;else surface.scrollLeft=0;return;}
 const expand=event.target.closest('[data-sheet-expand]');
 if(expand){const axis=expand.dataset.sheetExpand,start=Number(expand.dataset.start),end=Math.min(Number(expand.dataset.end),start+(axis==='rows'?49:11));excel[expand.dataset.side][axis==='rows'?'expandedRows':'expandedColumns'].push([start,end]);renderPreviews();}
});
$('show-highlights').addEventListener('change',event=>{state.showHighlights=event.target.checked;renderPreviews();});
$('page-select').addEventListener('change',event=>{if(state.result?.changeSet.format==='xlsx')selectPage(event.target.value);else{state.page=event.target.value;state.selected=state.result?.changeSet.format==='pdf'?state.result.changeSet.changes.find(c=>pdfPageKey(state.result,c)===state.page)?.id??null:null;renderChanges();renderPreviews({preserve:false,locate:!!state.selected});renderStructure();}});
function setChangesVisible(visible){$('changes-panel').hidden=!visible;$('workbench').classList.toggle('changes-collapsed',!visible);$('toggle-changes').setAttribute('aria-expanded',String(visible));$('toggle-changes').textContent=visible?'收起变更列表':'展开变更列表';}
$('toggle-changes').onclick=()=>setChangesVisible($('changes-panel').hidden);
document.addEventListener('click',event=>{if(!event.target.closest('#more-tools'))$('more-tools').open=false;});
document.addEventListener('keydown',event=>{if(event.key==='Escape'&&$('more-tools').open){$('more-tools').open=false;$('more-tools').querySelector('summary').focus();}});
// Opening help separately preserves the selected local files and current result.
document.querySelectorAll('a[href^="help.html"]').forEach(link=>{link.target='_blank';link.rel='noopener';});
$('export-json').onclick=()=>state.result&&blobDownload(JSON.stringify(state.result.changeSet,null,2),state.result.changeSet.format==='pdf'?'pdf-change.v1.json':'office-change.v1.json','application/json');
$('copy-json').onclick=async()=>{try{await navigator.clipboard.writeText(JSON.stringify(state.result.changeSet,null,2));toast('ChangeSet 已复制');}catch{toast('浏览器未允许复制，请使用导出按钮。');}};
async function compareClick(){if(state.busy)return;busy(true);error('');try{await runComparison();}catch(cause){clearComparison();finishProgress('对比失败',true);error(cause.message);}finally{busy(false);}}
$('compare').onclick=compareClick;
$('swap-files').onclick=async()=>{[state.files.before,state.files.after]=[state.files.after,state.files.before];setFiles();await compareClick();};
async function acceptFile(file,side){
 if(!file||state.busy)return;busy(true,'read');error('');
 try{
  const ext=file.name.split('.').pop().toLowerCase();if(!formats[ext])throw new Error('请选择 .docx、.xlsx、.pptx 或 .pdf 文件。');if(file.size>15*1024*1024)throw new Error('支持每个文件最大 15 MB。');
  state.files[side]={name:file.name,bytes:new Uint8Array(await file.arrayBuffer())};
  clearTimeout(finishProgress.timer);$('processing').hidden=true;clearComparison();setFiles();
  // Sample URLs should not reload an old example after custom files are selected.
  const url=new URL(location.href);if(url.searchParams.has('sample')){url.searchParams.delete('sample');history.replaceState(null,'',url);}
 }catch(cause){error(cause.message);}finally{busy(false);}
}
for(const [side,prefix] of [['before','old'],['after','new']]){
 $(prefix+'-input').addEventListener('change',async event=>{await acceptFile(event.target.files[0],side);event.target.value='';});$('replace-'+prefix).onclick=()=>$(prefix+'-input').click();
 const box=$('file-box-'+side);box.addEventListener('dragover',event=>{event.preventDefault();box.classList.add('drag-over');});box.addEventListener('dragleave',()=>box.classList.remove('drag-over'));box.addEventListener('drop',event=>{event.preventDefault();box.classList.remove('drag-over');acceptFile(event.dataTransfer.files[0],side);});
}
async function exportAnnotated(side){
 if(state.busy||!state.result)return;busy(true,'export');error('');
 try{await nextPaint();const engine=await engineReady;const result=await engine.annotate(state.files[side].bytes,state.result.changeSet,side);const file=state.files[side],ext=state.result.changeSet.format;blobDownload(result,file.name.replace(/\.[^.]+$/,'')+'.highlighted.'+ext,formats[ext].mime);toast('高亮副本已生成，原文件内容已保留。');}catch(cause){error(cause.message);}finally{busy(false);}
}
$('export-before').onclick=()=>exportAnnotated('before');$('export-after').onclick=()=>exportAnnotated('after');
const requestedSample=new URLSearchParams(location.search).get('sample');
if(requestedSample&&Object.hasOwn(samples,requestedSample))loadSample(requestedSample);
else if(requestedSample)error('未找到该示例，请从“示例与帮助”重新选择。');
