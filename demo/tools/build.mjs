import {readFile,writeFile,mkdir,cp,rm,readdir,stat} from 'node:fs/promises';
import {fileURLToPath} from 'node:url';
import path from 'node:path';
import wabtFactory from 'wabt';
import {createFormulaWorkbook} from './formula-fixture.mjs';
import {createPdfFixture} from './pdf-fixture.mjs';
const root=fileURLToPath(new URL('../',import.meta.url));
const web=path.join(root,'browser','wwwroot');
for(const side of ['old','new'])await writeFile(path.join(web,'samples',`excel-formulas-${side}.xlsx`),createFormulaWorkbook(side==='new'));
for(const side of ['old','new'])await writeFile(path.join(web,'samples',`pdf-budget-${side}.pdf`),await createPdfFixture(side==='new'));
const wabt=await wabtFactory();
const module=wabt.parseWat('character-diff.wat',await readFile(path.join(root,'wasm','character-diff.wat'),'utf8'));
module.resolveNames();module.validate();
const {buffer}=module.toBinary({canonicalize_lebs:true,write_debug_names:false});
await writeFile(path.join(web,'character-diff.wasm'),buffer);module.destroy();
await mkdir(path.join(web,'vendor'),{recursive:true});
await cp(path.join(root,'node_modules','fflate','esm','browser.js'),path.join(web,'vendor','fflate.js'));
await cp(path.join(root,'node_modules','fflate','LICENSE'),path.join(web,'vendor','fflate.LICENSE'));
const pdfVendor=path.join(web,'vendor','pdfjs');
await mkdir(pdfVendor,{recursive:true});
for(const name of ['pdf.mjs','pdf.worker.mjs'])await cp(path.join(root,'node_modules','pdfjs-dist','legacy','build',name),path.join(pdfVendor,name));
for(const name of ['cmaps','standard_fonts','wasm'])await cp(path.join(root,'node_modules','pdfjs-dist',name),path.join(pdfVendor,name),{recursive:true});
await cp(path.join(root,'node_modules','pdfjs-dist','LICENSE'),path.join(pdfVendor,'LICENSE'));
await cp(path.join(root,'node_modules','pdf-lib','dist','pdf-lib.esm.min.js'),path.join(web,'vendor','pdf-lib.js'));
await cp(path.join(root,'node_modules','pdf-lib','LICENSE.md'),path.join(web,'vendor','pdf-lib.LICENSE'));
for(const name of ['office-change.v1.schema.json','office-text-diff.v1.schema.json','office-impact.v1.schema.json','pdf-change.v1.schema.json']) {
  try {await cp(path.join(root,'..','schemas',name),path.join(web,'protocol',name));} catch(error) {if(error.code!=='ENOENT')throw error;}
}
// The exact output path is fixed and checked before replacing old generated assets.
const out=path.resolve(root,'out');
if(path.dirname(out)!==path.resolve(root)||path.basename(out)!=='out')throw new Error('Unsafe output path');
await rm(out,{recursive:true,force:true});
await cp(web,out,{recursive:true});
let bytes=0,files=0;
async function inventory(folder){for(const item of await readdir(folder,{withFileTypes:true})){const name=path.join(folder,item.name);if(item.isDirectory())await inventory(name);else{if(/\.(cs|c|dll|pdb|csproj)$/i.test(name)||name.includes('_framework'))throw new Error('Unexpected managed artifact: '+name);files++;bytes+=(await stat(name)).size;}}}
await inventory(out);
console.log(`Static frontend ready: ${files} files, ${bytes} bytes; standalone WASM ${buffer.length} bytes.`);
