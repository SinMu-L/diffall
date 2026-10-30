import {attr,nsAttr,number,children,first,elements,descendants,ancestors,is,elementPath,source,wordRuns,wordText,drawingText,sha256} from './ooxml.js';
const value=(type,data)=>({type,data});
const node=properties=>({key:'',parentKey:'root',pageKey:'',kind:'',label:'',index:0,text:null,scalar:null,formula:null,display:null,bold:false,fontSize:18,color:'18293D',fill:'transparent',geometry:null,asset:null,imageData:null,table:null,row:0,column:0,source:{},strongKey:true,...properties});
const view=(format,name)=>({format,name,pages:[],nodes:[]});
export function extractWord(pkg,name){
  const part=pkg.main(),root=pkg.xml(part),result=view('docx',name);result.pages.push({key:'body',name:'正文',width:1280,height:720});
  for(const p of descendants(root,'w','p').filter(p=>!ancestors(p).some(parent=>is(parent,'w','del')||is(parent,'w','moveFrom')))){
    const runs=wordRuns(p),text=wordText(p),id=Array.from(p.attributes).find(a=>a.localName==='paraId')?.value,pr=first(runs[0],'w','rPr'),bold=first(pr,'w','b'),index=result.nodes.length;
    result.nodes.push(node({key:'paragraph:'+(id??'position-'+index),strongKey:id!==undefined,pageKey:'body',kind:'paragraph',index,text,bold:!!bold&&!['0','false'].includes(nsAttr(bold,'w','val')),fontSize:Number(nsAttr(first(pr,'w','sz'),'w','val',24))/2,color:nsAttr(first(pr,'w','color'),'w','val','18293D'),label:text.length>25?text.slice(0,25)+'…':text||'空段落',source:source(part,{type:'word',elementPath:elementPath(p)})}));
  }return result;
}
export function extractExcel(pkg,name){
  const part=pkg.main(),root=pkg.xml(part),result=view('xlsx',name),sst=pkg.byType(part,'sharedStrings'),styles=pkg.byType(part,'styles');
  const strings=sst?elements(pkg.xml(sst),'s','si').map(si=>descendants(si,'s','t').map(t=>t.textContent).join('')):[];
  const style=styles?pkg.xml(styles):null,fonts=children(first(style,'s','fonts')),xfs=children(first(style,'s','cellXfs'));
  for(const sheet of descendants(root,'s','sheet')){
    const sheetId=number(sheet,'sheetId',1),sheetName=attr(sheet,'name'),sheetPart=pkg.related(part,nsAttr(sheet,'r','id'));if(!sheetPart)continue;
    const pageKey='sheet:'+sheetId;result.pages.push({key:pageKey,name:sheetName,width:1280,height:720,source:source(sheetPart,{type:'excel',sheetId})});
    const cells=descendants(pkg.xml(sheetPart),'s','c');
    const cellValue=cell=>{const type=attr(cell,'t'),raw=first(cell,'s','v')?.textContent;if(type==='inlineStr')return value('text',descendants(cell,'s','t').map(t=>t.textContent).join(''));if(type==='s'){const i=Number(raw);if(raw===undefined||!Number.isInteger(i)||i<0||i>=strings.length)throw new Error('共享字符串索引无效。');return value('text',strings[i]);}if(raw===undefined)return value('null',null);return type==='b'?value('boolean',raw==='1'):value(type==='e'?'error':['str','d'].includes(type)?'text':'number',raw);};
    const values=new Map(cells.map(cell=>[attr(cell,'r'),cellValue(cell)])),labelAt=address=>values.get(address)?.type==='text'?String(values.get(address).data):'';
    for(const cell of cells){const address=attr(cell,'r');if(!/^[A-Z]+[1-9][0-9]*$/.test(address))continue;const letters=address.match(/^[A-Z]+/)[0],row=Number(address.slice(letters.length));let column=0;for(const letter of letters)column=column*26+letter.charCodeAt(0)-64;
      const header=labelAt(letters+'1'),rowKey=labelAt('A'+row),xf=xfs[number(cell,'s')],font=fonts[number(xf,'fontId')],v=cellValue(cell),formula=first(cell,'s','f');
      result.nodes.push(node({key:pageKey+':'+(row===1&&header&&!formula?'header:'+header:row>1&&rowKey&&header?rowKey+':'+header:address),parentKey:pageKey,pageKey,kind:'cell',row,column,label:sheetName+' / '+(rowKey||'第'+row+'行')+' / '+(header||address),scalar:formula?value('null',null):v,cachedValue:formula&&first(cell,'s','v')?v:null,formulaType:formula?attr(formula,'t','normal'):null,display:String(v.data??''),formula:formula?.textContent??null,bold:!!first(font,'s','b')&&!['0','false'].includes(attr(first(font,'s','b'),'val')),fontSize:number(first(font,'s','sz'),'val',11),color:attr(first(font,'s','color'),'rgb','FF18293D').slice(-6),source:source(sheetPart,{type:'excel',sheetId,address})}));
    }
  }
  const groups=new Map();for(const n of result.nodes){if(!groups.has(n.key))groups.set(n.key,[]);groups.get(n.key).push(n);}for(const group of groups.values())if(group.length>1)for(const n of group)n.key=n.pageKey+':address:'+n.source.locator.address;
  return result;
}
export async function extractPowerPoint(pkg,name){
  const part=pkg.main(),root=pkg.xml(part),result=view('pptx',name),size=first(root,'p','sldSz'),width=number(size,'cx',12192000),height=number(size,'cy',6858000);let index=0;
  for(const ref of descendants(root,'p','sldId')){
    const id=number(ref,'id',index+256),slidePart=pkg.related(part,nsAttr(ref,'r','id'));if(!slidePart)throw new Error('幻灯片关系失效。');const slide=pkg.xml(slidePart),key='slide:'+id,title=descendants(slide,'a','t')[0]?.textContent??'';
    result.pages.push({key,name:title||'幻灯片 '+(index+1),width,height});result.nodes.push(node({key,pageKey:key,kind:'slide',index:index++,label:title||'幻灯片',source:source(slidePart,{type:'powerpoint',slideId:id})}));
    let shapeIndex=0;for(const shape of Array.from(slide.getElementsByTagName('*')).filter(el=>['sp','pic','graphicFrame'].some(name=>is(el,'p',name)))){
      const native=descendants(shape,'p','cNvPr')[0];if(!native)continue;const shapeId=number(native,'id'),path=ancestors(shape).filter(el=>is(el,'p','grpSp')).reverse().map(group=>number(descendants(group,'p','cNvPr')[0],'id')).concat(shapeId);
      const transform=first(first(shape,'p','spPr'),'a','xfrm')??first(shape,'p','xfrm'),off=first(transform,'a','off'),ext=first(transform,'a','ext'),body=first(shape,'p','txBody'),pr=descendants(body,'a','rPr')[0],table=descendants(shape,'a','tbl')[0],kind=is(shape,'p','pic')?'image':table?'table':'shape';
      const n=node({key:key+':shape:'+shapeId,parentKey:key,pageKey:key,index:shapeIndex++,kind,text:body?drawingText(body):null,label:title+' / '+attr(native,'name','元素'+shapeId),bold:attr(pr,'b')==='1',fontSize:number(pr,'sz',1800)/100,color:attr(first(first(pr,'a','solidFill'),'a','srgbClr'),'val','18293D'),fill:attr(first(first(first(shape,'p','spPr'),'a','solidFill'),'a','srgbClr'),'val','transparent'),geometry:{unit:'emu',space:path.length>1?'parent':'slide',x:number(off,'x'),y:number(off,'y'),width:number(ext,'cx'),height:number(ext,'cy'),rotation:number(transform,'rot'),flipH:attr(transform,'flipH')==='1',flipV:attr(transform,'flipV')==='1'},source:source(slidePart,{type:'powerpoint',slideId:id,shapePath:path})});
      if(kind==='image'){const imagePart=pkg.related(slidePart,nsAttr(descendants(shape,'a','blip')[0],'r','embed'));if(imagePart&&pkg.files[imagePart]){const bytes=pkg.files[imagePart],mediaType=pkg.contentType(imagePart);n.asset={sha256:await sha256(bytes),mediaType,part:'/'+imagePart};if(['image/png','image/jpeg','image/gif','image/webp'].includes(mediaType)){let binary='';for(let i=0;i<bytes.length;i+=32768)binary+=String.fromCharCode(...bytes.subarray(i,i+32768));n.imageData='data:'+mediaType+';base64,'+btoa(binary);}}}
      if(table){n.table=elements(table,'a','tr').map(row=>elements(row,'a','tc').map(cell=>drawingText(first(cell,'a','txBody'))));for(let r=0;r<n.table.length;r++)for(let c=0;c<n.table[r].length;c++)result.nodes.push(node({key:n.key+':cell:'+r+':'+c,parentKey:n.key,pageKey:key,kind:'cell',text:n.table[r][c],label:title+' / 表格 '+(r+1)+'行 '+(c+1)+'列',fontSize:16,source:source(slidePart,{type:'powerpoint',slideId:id,shapePath:path,cell:{row:r,column:c}})}));}
      result.nodes.push(n);
    }
  }return result;
}
