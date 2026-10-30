// Test-only server. It exposes the skill and synthetic project fixtures, not the workspace.
import http from 'node:http';
import {readFile,writeFile,readdir,mkdir} from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import {fileURLToPath} from 'node:url';

const root=fileURLToPath(new URL('../',import.meta.url));
const skill=path.join(root,'skills','office-diff');
const validator=process.argv[2]??path.join(os.homedir(),'.codex','skills','.system','skill-creator','scripts','quick_validate.py');
const pyodideRoot=path.join(root,'.tools','skill-smoke','pyodide','package');
const yamlRoot=path.join(root,'.tools','skill-validation','yaml');
const yamlFiles=[];
async function yamlInventory(folder){for(const entry of await readdir(folder,{withFileTypes:true})){const full=path.join(folder,entry.name);if(entry.isDirectory()){if(entry.name!=='__pycache__')await yamlInventory(full);}else if(entry.name.endsWith('.py'))yamlFiles.push(path.relative(yamlRoot,full).replaceAll('\\','/'));}}
await yamlInventory(yamlRoot);
const output=path.join(root,'demo','artifacts','skill-browser');
const files=[];
async function inventory(folder){for(const entry of await readdir(folder,{withFileTypes:true})){const full=path.join(folder,entry.name);if(entry.isDirectory()){if(entry.name!=='__pycache__')await inventory(full);}else if(!entry.name.endsWith('.pyc'))files.push(path.relative(skill,full).replaceAll('\\','/'));}}
await inventory(skill);
const samples=[['word-budget','docx',3],['excel-budget','xlsx',13],['ppt-budget','pptx',7],['excel-formulas','xlsx',null]];
const python=`
import runpy, json
from pathlib import Path
runtime = runpy.run_path('/skills/office-diff/scripts/office_diff.py', run_name='office_diff_skill')
self_test = await runtime['self_test']()
cases = []
report_html = None
for name, fmt, expected in ${JSON.stringify(samples).replaceAll('null','None')}:
    before = Path('/work/' + name + '-old.' + fmt)
    after = Path('/work/' + name + '-new.' + fmt)
    before_bytes, after_bytes = before.read_bytes(), after.read_bytes()
    result = await runtime['compare_files'](str(before), str(after), '/results/' + name, limit=2, max_chars=6000)
    full = json.loads(Path(result['artifacts']['changes']).read_text())
    if expected is not None:
        assert result['totalChanges'] == expected, (name, result['totalChanges'])
    assert full['coverage'] == result['coverage']
    assert len(json.dumps(result, ensure_ascii=False, separators=(',', ':'))) <= 6000
    assert before.read_bytes() == before_bytes and after.read_bytes() == after_bytes
    ids = [c['id'] for c in result['changes']]
    while result['hasMore']:
        result = runtime['query_changes']('/results/' + name + '/changes.json', offset=result['nextOffset'], limit=2, max_chars=6000)
        ids.extend(c['id'] for c in result['changes'])
    assert ids == [c['id'] for c in full['changes']]
    cases.append({'fixture': name, 'format': fmt, 'changes': len(full['changes']), 'pagesComplete': True, 'inputsUnchanged': True})
    if name == 'excel-budget':
        report_html = Path('/results/' + name + '/review.html').read_text()
validator = runpy.run_path('/validator.py', run_name='skill_validator')
valid, message = validator['validate_skill']('/skills/office-diff')
assert valid, message
browser_result = {'ok': True, 'runtime': 'Pyodide Web Worker', 'selfTest': self_test, 'cases': cases, 'skillValidation': message, 'runtimeNetworkDisabled': True}
json.dumps({'result': browser_result, 'report': report_html}, ensure_ascii=False)
`;
const worker=`self.onmessage=async()=>{try{
postMessage({phase:'正在加载测试宿主 Pyodide（技能包自身不下载依赖）'});
const {loadPyodide}=await import('/runtime/pyodide.mjs');
const py=await loadPyodide({indexURL:self.location.origin+'/runtime/'});
for(const name of await(await fetch('/validation-files')).json()){const destination='/validation/yaml/'+name;py.FS.mkdirTree(destination.slice(0,destination.lastIndexOf('/')));py.FS.writeFile(destination,new Uint8Array(await(await fetch('/validation/'+name)).arrayBuffer()));}
py.runPython("import sys; sys.path.insert(0, '/validation')");
const names=await(await fetch('/skill-files')).json();
for(const name of names){const destination='/skills/office-diff/'+name;py.FS.mkdirTree(destination.slice(0,destination.lastIndexOf('/')));py.FS.writeFile(destination,new Uint8Array(await(await fetch('/skill/'+name)).arrayBuffer()));}
py.FS.mkdirTree('/work');
for(const [name,format] of ${JSON.stringify(samples)})for(const side of ['old','new']){const file=name+'-'+side+'.'+format;py.FS.writeFile('/work/'+file,new Uint8Array(await(await fetch('/sample/'+file)).arrayBuffer()));}
py.FS.writeFile('/validator.py',await(await fetch('/validator.py')).text());
postMessage({phase:'正在浏览器 Worker 中运行原引擎、Python 桥接和结果分页'});
globalThis.fetch=()=>Promise.reject(new Error('Unexpected runtime network request'));
const value=JSON.parse(await py.runPythonAsync(${JSON.stringify(python)}));
postMessage(value);
}catch(error){postMessage({error:String(error.stack||error)});}};`;
const html=`<!doctype html><html lang="zh-CN"><meta charset="utf-8"><title>Office Skill 浏览器验证</title><style>body{font:16px/1.6 system-ui;padding:30px;max-width:1000px;margin:auto;background:#f7f9fb;color:#24384a}button{padding:10px 18px}pre{white-space:pre-wrap;overflow-wrap:anywhere;background:white;padding:20px;border:1px solid #dde4eb;border-radius:10px}</style><h1>Office Skill 浏览器验证</h1><p>测试仅使用项目中的合成文件。加载 Pyodide 后禁用 Worker 网络，再执行技能。</p><button id="run">运行浏览器验证</button><pre id="status">等待开始</pre><p><a href="/report" target="_blank">查看生成的 Excel 审阅报告</a></p><script>
document.querySelector('#run').onclick=()=>{document.querySelector('#run').disabled=true;const worker=new Worker('/worker.js',{type:'module'});worker.onmessage=async({data})=>{if(data.result){await fetch('/result',{method:'POST',body:JSON.stringify(data)});document.querySelector('#status').textContent=JSON.stringify(data.result,null,2);worker.terminate();}else if(data.error){document.querySelector('#status').textContent='失败：'+data.error;worker.terminate();}else document.querySelector('#status').textContent=data.phase;};worker.onerror=e=>{document.querySelector('#status').textContent='失败：'+e.message;};worker.postMessage({});};
</script></html>`;
const server=http.createServer(async(req,res)=>{try{
  const url=new URL(req.url,'http://127.0.0.1');
  const send=(type,data)=>{res.writeHead(200,{'Content-Type':type,'Cache-Control':'no-store'});res.end(data);};
  if(url.pathname==='/')return send('text/html; charset=utf-8',html);
  if(url.pathname==='/worker.js')return send('text/javascript; charset=utf-8',worker);
  if(url.pathname==='/validation-files')return send('application/json',JSON.stringify(yamlFiles));
  if(url.pathname.startsWith('/validation/')){const name=decodeURIComponent(url.pathname.slice(12));if(yamlFiles.includes(name))return send('text/plain; charset=utf-8',await readFile(path.join(yamlRoot,name)));}
  if(url.pathname.startsWith('/runtime/')){const name=decodeURIComponent(url.pathname.slice(9));if(/^[a-zA-Z0-9_.-]+$/.test(name)&&!name.includes('..'))return send(name.endsWith('.wasm')?'application/wasm':name.endsWith('.zip')?'application/zip':name.endsWith('.json')?'application/json':'text/javascript; charset=utf-8',await readFile(path.join(pyodideRoot,name)));}
  if(url.pathname==='/skill-files')return send('application/json',JSON.stringify(files));
  if(url.pathname==='/validator.py')return send('text/plain; charset=utf-8',await readFile(validator));
  if(url.pathname.startsWith('/skill/')){const name=decodeURIComponent(url.pathname.slice(7));if(files.includes(name))return send('text/plain; charset=utf-8',await readFile(path.join(skill,name)));}
  if(url.pathname.startsWith('/sample/')){const name=decodeURIComponent(url.pathname.slice(8));if(samples.some(([base,ext])=>['old','new'].some(side=>name===base+'-'+side+'.'+ext)))return send('application/octet-stream',await readFile(path.join(root,'demo','browser','wwwroot','samples',name)));}
  if(url.pathname==='/result'&&req.method==='POST'){let body='';for await(const chunk of req){body+=chunk;if(body.length>5e6)throw new Error('Oversize test output');}const result=JSON.parse(body);await mkdir(output,{recursive:true});await writeFile(path.join(output,'result.json'),JSON.stringify(result.result,null,2));await writeFile(path.join(output,'review.html'),result.report);return send('application/json','{"saved":true}');}
  if(url.pathname==='/report')return send('text/html; charset=utf-8',await readFile(path.join(output,'review.html')));
  res.writeHead(404);res.end('Not found');
}catch(error){res.writeHead(500,{'Content-Type':'text/plain; charset=utf-8'});res.end(String(error));}});
server.listen(4178,'127.0.0.1',()=>console.log('Office skill smoke test: http://127.0.0.1:4178/'));
