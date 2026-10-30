import {sourceKey} from './spreadsheet-review.js';
import {changeGroup} from './change-types.js';

export function pdfPageKey(result,change){
 for(const side of ['after','before']){const index=change?.[side]?.source?.locator?.pageIndex;if(index!=null){const page=result[side].pages.find(p=>p.index===index);if(page)return page.key;}}
 return null;
}
export async function renderPdfPreview(surface,{view,pageKey,side,changes,selected,showHighlights,engine}){
 const page=view.pages.find(p=>p.key===pageKey);
 if(!page){surface.innerHTML='<div class="preview-placeholder"><strong>此版本没有该页面</strong><p>该页在另一个版本中新增或删除。</p></div>';return;}
 const wrapper=document.createElement('div');wrapper.className='pdf-page';
 const width=Math.max(320,Math.min(900,surface.clientWidth-32));wrapper.style.width=width+'px';wrapper.style.aspectRatio=`${page.width} / ${page.height}`;
 const canvas=document.createElement('canvas');canvas.setAttribute('role','img');canvas.setAttribute('aria-label',`${side==='before'?'旧':'新'}版 PDF 第 ${page.index+1} 页`);wrapper.append(canvas);
 const byRegion=new Map();
 for(const change of changes){const loc=change[side]?.source?.locator;if(loc?.pageIndex!==page.index||(!showHighlights&&change.id!==selected))continue;
  const key=sourceKey(change[side].source),old=byRegion.get(key);if(!old||change.id===selected)byRegion.set(key,change);
 }
 for(const change of byRegion.values()){
  const rect=change[side].source.locator.rect,button=document.createElement('button');button.type='button';button.dataset.change=change.id;button.className=`pdf-highlight ${side}${change.id===selected?' selected':''}${change.target.kind==='page'?' pdf-whole-page':''}`;
  button.dataset.changeGroup=changeGroup(change);
  button.style.left=Math.max(0,rect[0]/page.width*100)+'%';button.style.top=Math.max(0,rect[1]/page.height*100)+'%';button.style.width=Math.min(100,rect[2]/page.width*100)+'%';button.style.height=Math.min(100,rect[3]/page.height*100)+'%';
  button.setAttribute('aria-label',change.target.label+' · '+({text:'文字修改',style:'样式变化',geometry:'位置或尺寸变化',existence:change.operation==='add'?'新增':'删除',order:'页面调序',parent:'跨页移动'}[change.aspect]||'变化'));button.title=button.getAttribute('aria-label');wrapper.append(button);
 }
 surface.replaceChildren(wrapper);
 const resolution=Math.min(2,(globalThis.devicePixelRatio||1)*width/page.width,Math.sqrt(8000000/(page.width*page.height)));
 await engine.render(view,page.index,canvas,resolution);
}
