const http = require('http');
const fs = require('fs');
const fsp = fs.promises;
const path = require('path');
const os = require('os');
const { spawn } = require('child_process');

const root = path.resolve(__dirname, '..');
const port = Number(process.env.PORT || 10000);
const maxDocBytes = 50 * 1024 * 1024;
const mime = {'.html':'text/html; charset=utf-8','.css':'text/css; charset=utf-8','.js':'application/javascript; charset=utf-8','.pdf':'application/pdf','.png':'image/png','.jpg':'image/jpeg','.jpeg':'image/jpeg','.webp':'image/webp','.json':'application/json; charset=utf-8'};

function json(res,status,body){res.writeHead(status,{'Content-Type':'application/json; charset=utf-8','Cache-Control':'no-store'});res.end(JSON.stringify(body));}
function collect(req){return new Promise((resolve,reject)=>{const chunks=[];let size=0;req.on('data',chunk=>{size+=chunk.length;if(size>maxDocBytes){reject(new Error('文件超过 50MB'));req.destroy();return;}chunks.push(chunk);});req.on('end',()=>resolve(Buffer.concat(chunks)));req.on('error',reject);});}
function runSoffice(input,dir){return new Promise((resolve,reject)=>{const command=process.env.SOFFICE_PATH||'soffice',profile=`file://${path.join(dir,'profile')}`;const child=spawn(command,[`-env:UserInstallation=${profile}`,'--headless','--convert-to','pdf','--outdir',dir,input],{stdio:['ignore','pipe','pipe']});let stderr='';child.stderr.on('data',d=>stderr+=d);const timer=setTimeout(()=>{child.kill('SIGKILL');reject(new Error('Word 转换超时'));},90000);child.on('error',err=>{clearTimeout(timer);reject(err);});child.on('close',code=>{clearTimeout(timer);code===0?resolve():reject(new Error(stderr||`转换程序退出：${code}`));});});}

async function convertWord(req,res,url){
  const requested=path.basename(url.searchParams.get('filename')||'document.docx');
  const match=requested.match(/\.(docx|doc)$/i);if(!match)return json(res,400,{error:'只接受 DOCX 或 DOC 文件'});
  const extension=match[1].toLowerCase();
  if(Number(req.headers['content-length']||0)>maxDocBytes)return json(res,413,{error:'Word 文件不能超过 50MB'});
  let dir;
  try{
    const body=await collect(req);if(!body.length)return json(res,400,{error:'Word 文件为空'});
    const oldDoc=body.length>=8&&body.subarray(0,8).equals(Buffer.from([0xd0,0xcf,0x11,0xe0,0xa1,0xb1,0x1a,0xe1])),docx=body.length>=4&&body.subarray(0,2).equals(Buffer.from('PK'));
    if((extension==='doc'&&!oldDoc)||(extension==='docx'&&!docx))return json(res,400,{error:'Word 文件格式与扩展名不匹配或文件已损坏'});
    dir=await fsp.mkdtemp(path.join(os.tmpdir(),'seal-word-'));const input=path.join(dir,`document.${extension}`);await fsp.writeFile(input,body);await runSoffice(input,dir);
    const output=path.join(dir,'document.pdf');const pdf=await fsp.readFile(output);res.writeHead(200,{'Content-Type':'application/pdf','Content-Length':pdf.length,'Cache-Control':'no-store'});res.end(pdf);
  }catch(err){console.error('Word conversion failed:',err);if(!res.headersSent)json(res,500,{error:'Word 转换失败，请确认文件未损坏后重试'});}
  finally{if(dir)await fsp.rm(dir,{recursive:true,force:true}).catch(()=>{});}
}

async function serveStatic(req,res,url){let pathname=decodeURIComponent(url.pathname);if(pathname==='/')pathname='/index.html';const file=path.resolve(root,'.'+pathname);if(!file.startsWith(root+path.sep))return json(res,403,{error:'禁止访问'});try{const stat=await fsp.stat(file);if(!stat.isFile())throw new Error('not file');res.writeHead(200,{'Content-Type':mime[path.extname(file).toLowerCase()]||'application/octet-stream','Cache-Control':/index\.html$/.test(file)?'no-cache':'public, max-age=3600'});fs.createReadStream(file).pipe(res);}catch{json(res,404,{error:'页面不存在'});}}

http.createServer(async(req,res)=>{const url=new URL(req.url,'http://localhost');if(req.method==='GET'&&url.pathname==='/health')return json(res,200,{ok:true});if(req.method==='POST'&&(url.pathname==='/api/convert-word'||url.pathname==='/api/convert-doc'))return convertWord(req,res,url);if(req.method==='GET'||req.method==='HEAD')return serveStatic(req,res,url);json(res,405,{error:'不支持此操作'});}).listen(port,()=>console.log(`Company seal tool listening on ${port}`));
