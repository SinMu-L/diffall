import {unzipSync,zipSync,strFromU8,strToU8} from '../vendor/fflate.js';
export const NS={w:'http://schemas.openxmlformats.org/wordprocessingml/2006/main',a:'http://schemas.openxmlformats.org/drawingml/2006/main',p:'http://schemas.openxmlformats.org/presentationml/2006/main',s:'http://schemas.openxmlformats.org/spreadsheetml/2006/main',r:'http://schemas.openxmlformats.org/officeDocument/2006/relationships',rel:'http://schemas.openxmlformats.org/package/2006/relationships',ct:'http://schemas.openxmlformats.org/package/2006/content-types',xml:'http://www.w3.org/XML/1998/namespace'};
export const children=node=>Array.from(node?.childNodes??[]).filter(child=>child.nodeType===1);
export const is=(node,prefix,name)=>node?.namespaceURI===NS[prefix]&&node.localName===name;
export const elements=(node,prefix,name)=>children(node).filter(child=>is(child,prefix,name));
export const first=(node,prefix,name)=>elements(node,prefix,name)[0];
export const descendants=(node,prefix,name)=>Array.from(node?.getElementsByTagNameNS(NS[prefix],name)??[]);
export const attr=(node,name,fallback='')=>node?.hasAttribute(name)?node.getAttribute(name):fallback;
export const nsAttr=(node,prefix,name,fallback='')=>node?.hasAttributeNS(NS[prefix],name)?node.getAttributeNS(NS[prefix],name):fallback;
export const number=(node,name,fallback=0)=>node?.hasAttribute(name)?Number(node.getAttribute(name)):fallback;
export function ancestors(node){const result=[];for(let p=node?.parentNode;p?.nodeType===1;p=p.parentNode)result.push(p);return result;}
export function elementPath(node){return ancestors(node).reverse().map(parent=>children(parent).indexOf([node,...ancestors(node)].find(child=>child.parentNode===parent)));}
export function atPath(root,path){let node=root;for(const index of path){if(!Number.isInteger(index)||index<0)throw new Error('无效的 XML 锚点。');node=children(node)[index];if(!node)throw new Error('原文件 XML 锚点失效。');}return node;}
export function create(parent,prefix,name,attributes={},text){const node=(parent.ownerDocument??parent).createElementNS(NS[prefix],(prefix==='s'?'':prefix+':')+name);for(const [key,value]of Object.entries(attributes)){if(value===null||value===undefined)continue;const [p,n]=key.split(':');if(n)node.setAttributeNS(NS[p],key,String(value));else node.setAttribute(key,String(value));}if(text!==undefined)node.textContent=text;return node;}
export function append(parent,prefix,name,attributes={},text){const node=create(parent,prefix,name,attributes,text);parent.appendChild(node);return node;}
export function remove(node){node?.parentNode?.removeChild(node);}
export function replace(node,replacements){for(const replacement of replacements)node.parentNode.insertBefore(replacement,node);remove(node);}
export function ordered(parent,node,order){remove(node);const rank=order.indexOf(node.localName);const later=children(parent).find(child=>order.indexOf(child.localName)>rank);parent.insertBefore(node,later??null);return node;}
export function parseXml(text){if(/<!DOCTYPE|<!ENTITY/i.test(text))throw new Error('Office XML 不支持 DTD 或实体声明。');const doc=new DOMParser().parseFromString(text,'application/xml');if(!doc.documentElement||doc.getElementsByTagName('parsererror').length)throw new Error('Office XML 无效。');return doc;}
export async function sha256(bytes){return Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',bytes)),value=>value.toString(16).padStart(2,'0')).join('');}
function zipGuard(bytes){
  if(bytes.length>15*1024*1024)throw new Error('Demo 支持每个文件最大 15 MB。');
  const view=new DataView(bytes.buffer,bytes.byteOffset,bytes.byteLength);let end=-1;
  for(let i=bytes.length-22;i>=Math.max(0,bytes.length-65557);i--)if(view.getUint32(i,true)===0x06054b50&&i+22+view.getUint16(i+20,true)===bytes.length){end=i;break;}
  if(end<0)throw new Error('文件不是有效的 OOXML ZIP 包。');
  const count=view.getUint16(end+10,true),size=view.getUint32(end+12,true),start=view.getUint32(end+16,true);
  if(view.getUint16(end+4,true)||view.getUint16(end+6,true)||count===65535||start===0xffffffff)throw new Error('Demo 不支持多卷 ZIP 或 ZIP64。');
  if(count>5000||start+size>end)throw new Error('文件包超过 Demo 的处理范围。');
  let offset=start,total=0;const names=new Set();
  for(let i=0;i<count;i++){
    if(offset+46>end||view.getUint32(offset,true)!==0x02014b50)throw new Error('ZIP 目录无效。');
    if(view.getUint16(offset+8,true)&1)throw new Error('请先解密受密码保护的 Office 文件。');
    const length=view.getUint16(offset+28,true),extra=view.getUint16(offset+30,true),comment=view.getUint16(offset+32,true);
    if(offset+46+length+extra+comment>end)throw new Error('ZIP 条目越界。');
    const name=strFromU8(bytes.subarray(offset+46,offset+46+length));
    if(names.has(name)||name.startsWith('/')||name.includes('\\')||name.split('/').includes('..'))throw new Error('ZIP 包含重复或无效路径。');names.add(name);
    total+=view.getUint32(offset+24,true);if(total>100*1024*1024)throw new Error('文件展开后超过 100 MB。');
    offset+=46+length+extra+comment;
  }
  if(offset!==start+size)throw new Error('ZIP 目录长度不一致。');
}
export class OfficePackage {
  constructor(bytes){zipGuard(bytes);this.files=unzipSync(bytes);this.docs=new Map();this.dirty=new Set();if(!this.files['[Content_Types].xml'])throw new Error('缺少 OOXML 内容类型。');}
  xml(part){part=part.replace(/^\//,'');if(!this.docs.has(part)){if(!this.files[part])throw new Error('缺少 Office part：'+part);this.docs.set(part,parseXml(strFromU8(this.files[part])));}return this.docs.get(part).documentElement;}
  mark(part){this.dirty.add(part.replace(/^\//,''));}
  relationships(part=''){const folder=part.includes('/')?part.slice(0,part.lastIndexOf('/')+1):'',name=part.slice(part.lastIndexOf('/')+1);const rel=folder+'_rels/'+name+'.rels';if(!this.files[rel])return[];return descendants(this.xml(rel),'rel','Relationship');}
  resolve(part,target){const base=part.split('/');base.pop();for(const fragment of target.replace(/^\//,'').split('/')){if(fragment==='..'){if(!base.length)throw new Error('关系路径越界。');base.pop();}else if(fragment!=='.'&&fragment)base.push(fragment);}return base.join('/');}
  related(part,id){const rel=this.relationships(part).find(item=>attr(item,'Id')===id);if(!rel||attr(rel,'TargetMode')==='External')return null;return this.resolve(attr(rel,'Target').startsWith('/')?'':part,attr(rel,'Target'));}
  byType(part,type){const rel=this.relationships(part).find(item=>attr(item,'Type').endsWith('/'+type)&&attr(item,'TargetMode')!=='External');return rel?this.resolve(attr(rel,'Target').startsWith('/')?'':part,attr(rel,'Target')):null;}
  main(){const part=this.byType('','officeDocument');if(!part)throw new Error('缺少 Office 主文档关系。');return part;}
  contentType(part){const root=this.xml('[Content_Types].xml');return attr(descendants(root,'ct','Override').find(item=>attr(item,'PartName')==='/'+part),'ContentType')||attr(descendants(root,'ct','Default').find(item=>attr(item,'Extension')===part.split('.').pop()),'ContentType');}
  write(){if(!this.dirty.size)return null;for(const part of this.dirty)this.files[part]=strToU8(new XMLSerializer().serializeToString(this.docs.get(part)));return zipSync(this.files,{level:6});}
}
export const source=(part,locator)=>({part:'/'+part,locator});
export const wordRuns=p=>descendants(p,'w','r').filter(run=>ancestors(run).find(node=>is(node,'w','p'))===p&&!ancestors(run).some(node=>is(node,'w','del')||is(node,'w','moveFrom')));
export const wordChildText=child=>is(child,'w','t')?child.textContent:is(child,'w','tab')?'\t':is(child,'w','br')||is(child,'w','cr')?'\n':'';
export const wordText=p=>wordRuns(p).flatMap(run=>children(run).map(wordChildText)).join('');
export const drawingText=body=>elements(body,'a','p').map(p=>children(p).map(run=>is(run,'a','r')||is(run,'a','fld')?first(run,'a','t')?.textContent??'':is(run,'a','br')?'\n':'').join('')).join('\n');
