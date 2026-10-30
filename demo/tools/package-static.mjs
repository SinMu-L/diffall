import {readFile,writeFile,readdir,mkdir} from 'node:fs/promises';
import {fileURLToPath} from 'node:url';
import path from 'node:path';
import {zipSync,strToU8} from 'fflate';
const root=fileURLToPath(new URL('../',import.meta.url)),out=path.join(root,'out'),files={};
async function add(folder){for(const item of await readdir(folder,{withFileTypes:true})){const full=path.join(folder,item.name);if(item.isDirectory())await add(full);else files[path.relative(out,full).split(path.sep).join('/')]=new Uint8Array(await readFile(full));}}
await add(out);
files['START.txt']=strToU8('Office Diff Lab: static HTML/CSS/JS + standalone WebAssembly.\nServe this folder using any HTTP/HTTPS static server, then open index.html.\nNo C/C#, .NET, backend API, Node runtime or CDN required by the deployed app.\nDo not use file:// because browser ES modules and Web Crypto require HTTP(S).\n');
await mkdir(path.join(root,'artifacts'),{recursive:true});
const destination=path.join(root,'artifacts','office-diff-static.zip'),bytes=zipSync(files,{level:9});await writeFile(destination,bytes);console.log(`Static delivery: ${destination} (${bytes.length} bytes)`);
