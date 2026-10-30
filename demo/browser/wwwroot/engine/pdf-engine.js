import {sha256} from './ooxml.js';
import {comparePdfModels,groupSpans,round} from './pdf-model.js';
let runtime;
export async function pdfRuntime(){
 if(!runtime)runtime=import('../vendor/pdfjs/pdf.mjs').then(pdfjs=>{pdfjs.GlobalWorkerOptions.workerSrc=new URL('../vendor/pdfjs/pdf.worker.mjs',import.meta.url).href;return pdfjs;}).catch(error=>{runtime=null;throw error;});
 return runtime;
}
const assets=new URL('../vendor/pdfjs/',import.meta.url).href;
const assetPath=typeof process!=='undefined'&&process.versions?.node?decodeURIComponent(new URL('../vendor/pdfjs/',import.meta.url).pathname).replace(/^\/([A-Za-z]:)/,'$1'):assets;
async function openPdf(bytes){
 if(bytes.length>15*1024*1024)throw new Error('PDF 每个文件最大支持 15 MB。');
 if(!new TextDecoder().decode(bytes.subarray(0,1024)).includes('%PDF-'))throw new Error('文件内容不是有效的 PDF。');
 const pdfjs=await pdfRuntime();
 const task=pdfjs.getDocument({data:bytes.slice(),cMapUrl:assetPath+'cmaps/',cMapPacked:true,standardFontDataUrl:assetPath+'standard_fonts/',wasmUrl:assetPath+'wasm/',fontExtraProperties:true,isEvalSupported:false,stopAtErrors:true,useSystemFonts:false});
 try{const doc=await task.promise;if(doc.numPages>150)throw new Error('PDF 演示最多支持 150 页，请拆分后比较。');return {doc,task,pdfjs};}
 catch(error){await task.destroy();if(error.name==='PasswordException')throw new Error('PDF 受密码保护，请先解密后比较。');throw error;}
}
function fontStyle(page,fontName,size,styles){
 let font;try{font=page.commonObjs.get(fontName);}catch{}
 const name=font?.name?.replace(/^[A-Z]{6}\+/,'')||null;
 return {fontFamily:name?.replace(/[-,](?:Bold(?:Italic|Oblique)?|Italic|Oblique|Regular)(?:MT)?$/i,'')??null,fontSize:round(size),bold:font?Boolean(font.bold||/(?:Bold|Black|Heavy|Demi)/i.test(name)):null,italic:font?Boolean(font.italic||/(?:Italic|Oblique)/i.test(name)):null};
}
/** Color is assigned only when glyphs map exactly to extracted Unicode text. */
function paintedCharacters(list,OPS){
 let state={font:null,color:'#000000',mode:0},stack=[];const fonts=new Map();
 for(let i=0;i<list.fnArray.length;i++){
  const fn=list.fnArray[i],args=list.argsArray[i];
  if(fn===OPS.save||fn===OPS.paintFormXObjectBegin)stack.push({...state});
  else if(fn===OPS.restore||fn===OPS.paintFormXObjectEnd)state=stack.pop()||{font:null,color:null,mode:0};
  else if(fn===OPS.setFont)state.font=args[0];
  else if(fn===OPS.setFillRGBColor)state.color=/^#[0-9a-f]{6}$/i.test(args[0])?args[0].toLowerCase():null;
  else if(fn===OPS.setFillColorN||fn===OPS.setFillTransparent)state.color=null;
  else if(fn===OPS.setTextRenderingMode)state.mode=args[0];
  else if(fn===OPS.showText){if(!fonts.has(state.font))fonts.set(state.font,{text:'',colors:[],cursor:0});const record=fonts.get(state.font);
   for(const glyph of args[0])if(typeof glyph==='object'&&glyph?.unicode)for(const char of glyph.unicode){if(/\s/u.test(char))continue;record.text+=char;for(let n=0;n<char.length;n++)record.colors.push(state.mode===0?state.color:null);}
  }
 }
 return fonts;
}
function spanStyles(item,base,paint){
 const record=paint.get(item.fontName),text=item.str.replace(/\s/gu,'');
 const start=record?.text.indexOf(text,record.cursor)??-1,known=start>=0&&start-record.cursor<4;
 if(known)record.cursor=start+text.length;
 let offset=0;
 return Array.from({length:item.str.length},(_,i)=>{const color=/\s/u.test(item.str[i])?null:known?record.colors[start+offset]:null;if(!/\s/u.test(item.str[i]))offset++;return {...base,color};});
}
export class PdfEngine{
 constructor(differ){this.differ=differ;this.sessions=new Set();}
 async extract(bytes,name){
  const session=await openPdf(bytes),{doc,pdfjs}=session;const pages=[],issues=[];let count=0,characters=0;
  try{
   for(let index=0;index<doc.numPages;index++){
    const proxy=await doc.getPage(index+1),viewport=proxy.getViewport({scale:1});
    const content=await proxy.getTextContent({disableNormalization:true}),list=await proxy.getOperatorList();count+=content.items.length;characters+=content.items.reduce((n,item)=>n+(item.str?.length||0),0);
    if(count>50000||content.items.length>3000||characters>500000)throw new Error('PDF 文字对象超过演示处理上限，请拆分文件后比较。');
    const paint=paintedCharacters(list,pdfjs.OPS),page={kind:'page',key:'pdf:raw:'+index,index,name:`第 ${index+1} 页`,label:`第 ${index+1} 页`,width:round(viewport.width),height:round(viewport.height),rotation:proxy.rotate,geometry:{unit:'pt',space:'pageViewport',x:0,y:0,width:round(viewport.width),height:round(viewport.height),rotation:proxy.rotate},source:{locator:{type:'pdf',pageIndex:index,coordinateSpace:'pageViewport',rect:[0,0,round(viewport.width),round(viewport.height)]}}};
    const spans=[];
    for(const [itemIndex,item]of content.items.entries()){
     if(!item.str?.trim())continue;
     const style=content.styles[item.fontName],t=pdfjs.Util.transform(viewport.transform,item.transform),size=Math.hypot(t[2],t[3]),angle=Math.atan2(t[1],t[0]),ascent=style?.ascent??(style?.descent!=null?1+style.descent:.8),descent=style?.descent??-.2;
     // Transform a baseline rectangle into the rotated, cropped viewport.
     const ux=Math.cos(angle),uy=Math.sin(angle),vx=uy,vy=-ux,width=item.width*viewport.scale*proxy.userUnit;
     const points=[[t[4]+vx*size*ascent,t[5]+vy*size*ascent],[t[4]+ux*width+vx*size*ascent,t[5]+uy*width+vy*size*ascent],[t[4]+vx*size*descent,t[5]+vy*size*descent],[t[4]+ux*width+vx*size*descent,t[5]+uy*width+vy*size*descent]];
     const xs=points.map(p=>p[0]),ys=points.map(p=>p[1]);
     spans.push({text:item.str,itemIndex,size,angle:angle*180/Math.PI,baseline:[t[4],t[5]],rect:[Math.min(...xs),Math.min(...ys),Math.max(...xs)-Math.min(...xs),Math.max(...ys)-Math.min(...ys)],styles:spanStyles(item,fontStyle(proxy,item.fontName,size,content.styles),paint)});
    }
    page.lines=groupSpans(spans,page);page.text=page.lines.map(n=>n.text).join('\n');pages.push(page);
    if(!page.lines.length)issues.push({code:'opaqueContent',side:'both',message:`第 ${index+1} 页没有可用文字层，扫描图片或转曲文字未进行内容比较。`,source:page.source});
    if(page.lines.some(n=>n.styles.some((s,i)=>! /\s/u.test(n.text[i])&&(s.color===null||s.fontFamily===null))))issues.push({code:'unresolvedStyle',side:'both',message:`第 ${index+1} 页部分字体或颜色无法可靠提取，未推断其属性变化。`,source:page.source});
   }
   if(!pages.some(p=>p.lines.length))throw new Error('此 PDF 没有可提取的文字。当前版本支持带文字层的 PDF，扫描件 OCR 尚未接入。');
   const view={format:'pdf',name,pages,nodes:pages.flatMap(p=>[p,...p.lines]),issues};
   this.sessions.add(session);Object.defineProperty(view,'pdfSession',{value:session,enumerable:false});return view;
  }catch(error){await session.task.destroy();throw error;}
 }
 async dispose(view){const session=view?.pdfSession;if(session&&this.sessions.delete(session))await session.task.destroy();}
 async compare(oldBytes,newBytes,oldName,newName,{onProgress}={}){
  let before,after;
  try{
   await onProgress?.('before');before=await this.extract(oldBytes,oldName);
   await onProgress?.('after');after=await this.extract(newBytes,newName);
   await onProgress?.('diff');const result=comparePdfModels(before,after,this.differ);
   const [oldHash,newHash]=await Promise.all([sha256(oldBytes),sha256(newBytes)]);
   const issues=[{code:'unsupportedFeature',side:'both',message:'比较文字、可提取的字体/字号/粗体/斜体/填充颜色、文字位置、页面尺寸及页序。图片、图表、线条、背景、批注和 OCR 未比较；多栏、竖排、复杂变换、跨行重排与重复内容的匹配可能不完整。高亮按文本区域标记，明细提供字符差异。'},...before.issues.map(i=>({...i,side:'before'})),...after.issues.map(i=>({...i,side:'after'})),...result.issues];
   return {before,after,changeSet:{schemaVersion:'1.0.0',id:'compare-'+oldHash.slice(0,8)+'-'+newHash.slice(0,8),format:'pdf',documents:{before:{name:oldName,sha256:oldHash},after:{name:newName,sha256:newHash}},comparison:{profile:'pdf-text-v1',aspects:['existence','text','style','geometry','order','parent'],geometryTolerancePt:.75},coverage:{status:'partial',issues},changes:result.changes}};
  }catch(error){await this.dispose(before);await this.dispose(after);throw error;}
 }
 async render(view,pageIndex,canvas,scale=1){const page=await view.pdfSession.doc.getPage(pageIndex+1),viewport=page.getViewport({scale});canvas.width=Math.ceil(viewport.width);canvas.height=Math.ceil(viewport.height);const task=page.render({canvasContext:canvas.getContext('2d'),canvas,viewport});return task.promise;}
 async annotate(bytes,set,side){
  if(!['before','after'].includes(side))throw new Error('高亮版本无效。');
  if(set.format!=='pdf'||set.comparison?.profile!=='pdf-text-v1'||set.schemaVersion!=='1.0.0')throw new Error('不支持的 PDF 差异协议。');
  if(await sha256(bytes)!==set.documents[side].sha256)throw new Error('文件快照与差异结果不一致，请重新比较。');
  const {PDFDocument,rgb}=await import('../vendor/pdf-lib.js');
  const doc=await PDFDocument.load(bytes,{updateMetadata:false}),view=await this.extract(bytes,'snapshot.pdf'),session=view.pdfSession;
  try{
   const rects=new Set();
   for(const change of set.changes){const s=change[side];if(!s)continue;const loc=s.source?.locator;
    if(loc?.type!=='pdf'||loc.coordinateSpace!=='pageViewport'||!Number.isInteger(loc.pageIndex)||loc.pageIndex<0||loc.pageIndex>=doc.getPageCount()||!Array.isArray(loc.rect)||loc.rect.length!==4||!loc.rect.every(Number.isFinite)||loc.rect[2]<=0||loc.rect[3]<=0)throw new Error('PDF 定位锚点无效。');
    const proxy=await session.doc.getPage(loc.pageIndex+1),viewport=proxy.getViewport({scale:1}),[x,y,w,h]=loc.rect;
    if(x<-.5||y<-.5||x+w>viewport.width+1||y+h>viewport.height+1)throw new Error('PDF 区域超出页面边界，无法准确高亮。');
    const canonical=loc.text==null?view.pages[loc.pageIndex].source.locator:view.pages[loc.pageIndex].lines.find(n=>JSON.stringify(n.source.locator.itemIndices)===JSON.stringify(loc.itemIndices))?.source.locator;
    if(!canonical||loc.rect.some((v,i)=>Math.abs(v-canonical.rect[i])>.001))throw new Error('PDF 定位锚点与原始区域不一致。');
    if(loc.text!=null){if(loc.text!==canonical.text||JSON.stringify(loc.range)!==JSON.stringify(canonical.range))throw new Error('PDF 文字锚点与原文不一致。');
     if(change.aspect==='text'&&s.value.data!==loc.text)throw new Error('PDF 文字值与原文锚点不一致。');
    }
    const key=JSON.stringify(loc);if(rects.has(key))continue;rects.add(key);
    const points=[[x,y],[x+w,y],[x,y+h],[x+w,y+h]].map(p=>viewport.convertToPdfPoint(...p)),xs=points.map(p=>p[0]),ys=points.map(p=>p[1]);
    const page=doc.getPage(loc.pageIndex),color=side==='before'?rgb(.89,.57,.16):rgb(.17,.65,.44);
    page.drawRectangle({x:Math.min(...xs),y:Math.min(...ys),width:Math.max(...xs)-Math.min(...xs),height:Math.max(...ys)-Math.min(...ys),color,opacity:change.target.kind==='page'?.035:.18,borderColor:color,borderWidth:1,borderOpacity:.8});
   }
   return await doc.save({useObjectStreams:false});
  }finally{await this.dispose(view);}
 }
}
