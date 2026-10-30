import {PDFDocument,StandardFonts,rgb,degrees,PDFString,PDFName} from 'pdf-lib';

/** Deterministic fixtures exercise text, mixed styles, page insertion/removal and reorder. */
export async function createPdfFixture(revised=false,{rotation=0,crop=false}={}){
 const doc=await PDFDocument.create();doc.setCreationDate(new Date('2026-01-01T00:00:00Z'));doc.setModificationDate(new Date('2026-01-01T00:00:00Z'));
 const regular=await doc.embedFont(StandardFonts.Helvetica),bold=await doc.embedFont(StandardFonts.HelveticaBold);
 const page=(title,lines)=>{const p=doc.addPage([595,842]);p.drawRectangle({x:36,y:758,width:523,height:48,color:rgb(.08,.18,.23)});p.drawText(title,{x:52,y:775,size:21,font:bold,color:rgb(1,1,1)});for(const [text,y,options={}]of lines)p.drawText(text,{x:52,y,size:13,font:regular,color:rgb(.14,.21,.25),...options});return p;};
 const summary=()=>{const p=page('Budget review',[
  ['Project: North Star',714],['Approved budget: '+(revised?'120':'100')+' million',676],
  ['Delivery schedule',620,{size:revised?17:13}],['Status: ready for review',572,{color:revised?rgb(.7,.12,.15):rgb(.14,.21,.25)}],
  ['Owner: Operations',revised?500:526],['Revision notes',450]
 ]);p.drawText('Mixed style: ',{x:52,y:418,size:13,font:regular});p.drawText('priority',{x:128,y:418,size:13,font:revised?bold:regular});
 if(revised)p.drawText('New: monthly review checkpoints.',{x:52,y:374,size:13,font:regular});
 if(rotation)p.setRotation(degrees(rotation));if(crop)p.setCropBox(20,30,550,780);return p;};
 const risk=()=>page('Risk register',[['Risk A: supplier lead time',714],['Risk B: resource availability',676],['Mitigation: weekly follow-up',638]]);
 const approval=()=>page('Approval record',[['Sponsor: approved',714],['Finance: pending final review',676],['Next checkpoint: November',638]]);
 if(revised){page('New review checklist',[['Confirm owners and review dates.',714],['Share the approved budget summary.',676]]);risk();summary();}
 else{summary();risk();approval();}
 return doc.save({useObjectStreams:false});
}

/** Standard CJK CID font fixture exercises bundled CMaps without machine-specific fonts. */
export async function createCjkPdfFixture(revised=false){
 const doc=await PDFDocument.create(),ctx=doc.context;
 const descriptor=ctx.register(ctx.obj({Type:'FontDescriptor',FontName:'STSong-Light',Flags:6,FontBBox:[-25,-254,1000,880],ItalicAngle:0,Ascent:880,Descent:-120,CapHeight:880,StemV:80}));
 const descendant=ctx.register(ctx.obj({Type:'Font',Subtype:'CIDFontType0',BaseFont:'STSong-Light',CIDSystemInfo:{Registry:PDFString.of('Adobe'),Ordering:PDFString.of('GB1'),Supplement:4},FontDescriptor:descriptor,DW:1000}));
 const font=ctx.register(ctx.obj({Type:'Font',Subtype:'Type0',BaseFont:'STSong-Light',Encoding:'UniGB-UCS2-H',DescendantFonts:[descendant]}));
 const page=doc.addPage([595,842]);page.node.set(PDFName.of('Resources'),ctx.obj({Font:{CJK:font}}));
 const text='预算审阅：项目金额 '+(revised?'150':'120')+' 万元',hex=Array.from(text,c=>c.charCodeAt(0).toString(16).padStart(4,'0')).join('');
 page.node.addContentStream(ctx.register(ctx.flateStream(`BT /CJK 16 Tf 1 0 0 1 52 700 Tm <${hex}> Tj ET`)));
 return doc.save({useObjectStreams:false});
}
