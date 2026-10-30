import http from 'node:http';
import {readFile,stat} from 'node:fs/promises';
import {fileURLToPath} from 'node:url';
import path from 'node:path';
const root=path.join(path.dirname(fileURLToPath(import.meta.url)),'out');
const port=Number(process.env.PORT||4173);
const types={'.html':'text/html; charset=utf-8','.js':'text/javascript; charset=utf-8','.mjs':'text/javascript; charset=utf-8','.pdf':'application/pdf','.css':'text/css; charset=utf-8','.json':'application/json; charset=utf-8','.wasm':'application/wasm','.webmanifest':'application/manifest+json','.svg':'image/svg+xml'};
const server=http.createServer(async(req,res)=>{
 try{const pathname=decodeURIComponent(new URL(req.url,'http://localhost').pathname);const filename=path.resolve(root,'.'+(pathname.endsWith('/')?pathname+'index.html':pathname));if(filename!==root&&!filename.startsWith(root+path.sep)){res.writeHead(403).end();return;}const info=await stat(filename);if(!info.isFile())throw new Error('not a file');res.writeHead(200,{'Content-Type':types[path.extname(filename)]||'application/octet-stream','Cache-Control':'no-cache'});res.end(await readFile(filename));}catch{res.writeHead(404,{'Content-Type':'text/plain; charset=utf-8'}).end('Not found');}
});
server.listen(port,'127.0.0.1',()=>console.log(`Local: http://127.0.0.1:${port}`));
process.on('SIGINT',()=>server.close(()=>process.exit(0)));
