// Download an integrity-checked official Pyodide runtime for the local browser test.
// This runtime is a test dependency; it is never included in the skill archive.
import {mkdir,writeFile,readFile} from 'node:fs/promises';
import {createHash} from 'node:crypto';
import {execFileSync} from 'node:child_process';
import {fileURLToPath} from 'node:url';
import path from 'node:path';
const root=fileURLToPath(new URL('../',import.meta.url));
const target=path.join(root,'.tools','skill-smoke','pyodide');
const version='314.0.7';
const metadata=await(await fetch('https://registry.npmjs.org/pyodide/'+version)).json();
const response=await fetch(metadata.dist.tarball);
if(!response.ok)throw new Error('Pyodide download HTTP '+response.status);
const bytes=Buffer.from(await response.arrayBuffer());
const [algorithm,expected]=metadata.dist.integrity.split('-');
if(createHash(algorithm).update(bytes).digest('base64')!==expected)throw new Error('Pyodide package integrity mismatch');
await mkdir(target,{recursive:true});
const archive=path.join(target,'pyodide.tgz');await writeFile(archive,bytes);
const entries=execFileSync('tar',['-tzf',archive],{encoding:'utf8'}).trim().split(/\r?\n/);
if(entries.some(name=>!name.startsWith('package/')||name.includes('..')))throw new Error('Unexpected package paths');
execFileSync('tar',['-xzf',archive,'-C',target]);
await readFile(path.join(target,'package','pyodide.asm.wasm'));
console.log(JSON.stringify({version,downloadBytes:bytes.length,runtime:path.join(target,'package')}));
