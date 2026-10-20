const pdfjsLib = window.pdfjsLib;
if (!pdfjsLib || !window.PDFLib) throw new Error('PDF 组件未能载入');
pdfjsLib.GlobalWorkerOptions.workerSrc = './vendor/pdf.worker.min.js?v=23';

const $ = (id) => document.getElementById(id);
const state = { pdfBytes:null, pdfDoc:null, sealDataUrl:null, originalSealDataUrl:null, signatureDataUrl:null, signatureAspect:2.5, signatureSlots:[null,null], activeSignature:0, placements:[], signatures:[], selectedId:null, mode:'normal', edge:'right', seamInsetMm:5, fileName:'', startedAt:null };
const els = { pdfInput:$('pdfInput'), sealInput:$('sealInput'), signatureInput:$('signatureInput'), signatureInput2:$('signatureInput2'), pages:$('pages'), empty:$('emptyState'), controls:$('controls'), export:$('exportBtn'), manifest:$('manifestBtn'), auth:$('confirmAuth'), status:$('status') };

function toast(message){ const el=$('toast'); el.textContent=message; el.classList.add('show'); clearTimeout(toast.timer); toast.timer=setTimeout(()=>el.classList.remove('show'),3200); }
function ready(){ const ok=state.pdfBytes&&state.sealDataUrl; els.controls.classList.toggle('disabled',!ok); els.export.disabled=!(ok&&els.auth.checked); els.manifest.disabled=!state.pdfBytes; els.status.classList.remove('loading');els.status.classList.toggle('ready-mark',!!ok);els.status.textContent=!state.pdfBytes?'程序已就绪 · 请选择 PDF 或 Word':!state.sealDataUrl?'文件已载入 · 请上传印章图片':!els.auth.checked?'文件和印章已载入 · 请确认授权':'已就绪，所有处理均在本机完成'; }
function readFile(file, asDataUrl=false){ return new Promise((resolve,reject)=>{ const r=new FileReader(); r.onload=()=>resolve(r.result); r.onerror=reject; asDataUrl?r.readAsDataURL(file):r.readAsArrayBuffer(file); }); }

els.pdfInput.addEventListener('change',async e=>{ const file=e.target.files[0]; if(!file)return; try{$('pdfLabel').textContent='正在打开…';const raw=await readFile(file),isDocx=/\.docx$/i.test(file.name);let bytes;if(isDocx){$('pdfLabel').textContent='正在转换 Word 页面…';bytes=new Uint8Array(await convertDocxToPdf(raw));}else{bytes=new Uint8Array(raw);if(bytes.length<5||String.fromCharCode(...bytes.slice(0,5))!=='%PDF-')throw new Error('请选择有效的 PDF 或 DOCX 文件');}state.fileName=file.name;state.pdfBytes=bytes;state.startedAt=new Date().toISOString();state.placements=[];state.signatures=[];state.selectedId=null;await renderPdf();$('pdfLabel').textContent=file.name;$('docName').textContent=file.name;toast(`已打开 ${state.pdfDoc.numPages} 页${isDocx?' Word':' PDF'}`);}catch(err){console.error(err);state.pdfBytes=null;state.pdfDoc=null;$('pdfLabel').textContent='重新选择 PDF 或 Word';els.pages.innerHTML='';els.empty.hidden=false;toast('文件打开失败：'+friendlyError(err));}finally{e.target.value='';ready();} });
els.sealInput.addEventListener('change',async e=>{ const file=e.target.files[0]; if(!file)return; try{$('sealLabel').textContent='正在抠取印章…';state.originalSealDataUrl=await readFile(file,true);state.sealDataUrl=await optimizeSeal(state.originalSealDataUrl);$('sealLabel').textContent=file.name;$('sealPreview').src=state.sealDataUrl;$('sealPreview').hidden=false;$('sealPlaceholder').hidden=true;$('imageTools').hidden=false;refreshSealImages();toast('印章已抠取，背景已透明');}catch(err){console.error(err);state.sealDataUrl=null;$('sealLabel').textContent='重新选择印章图片';toast('图片处理失败：'+(err?.message||'请换一张常见格式图片'));}finally{e.target.value='';ready();} });
async function handleSignatureUpload(e,slot){const file=e.target.files[0];if(!file)return;const suffix=slot===0?'':'2';try{$('signatureLabel'+suffix).textContent=`正在处理签名 ${slot+1}…`;const source=await readFile(file,true),result=await optimizeSignature(source);state.signatureSlots[slot]=result;setActiveSignature(slot);$('signatureLabel'+suffix).textContent=file.name;$('signaturePreview'+suffix).src=result.dataUrl;$('signaturePreview'+suffix).hidden=false;$('signaturePlaceholder'+suffix).hidden=true;document.querySelectorAll(`.signature-mark[data-signature-slot="${slot}"] img`).forEach(i=>i.src=result.dataUrl);setMode('signature');toast(`签名 ${slot+1} 已准备好，请点击文件页面插入`);}catch(err){console.error(err);$('signatureLabel'+suffix).textContent=`重新选择签名 ${slot+1}`;toast('签名处理失败：'+(err?.message||'无法识别图片'));}finally{e.target.value='';}}
els.signatureInput.addEventListener('change',e=>handleSignatureUpload(e,0));
els.signatureInput2.addEventListener('change',e=>handleSignatureUpload(e,1));
els.auth.addEventListener('change',ready);

async function convertDocxToPdf(arrayBuffer){
  if(!window.docx?.renderAsync||!window.html2canvas)throw new Error('Word 转换组件未能载入，请刷新页面后重试');
  const host=document.createElement('div');host.className='docx-convert-host';document.body.append(host);
  try{
    await window.docx.renderAsync(arrayBuffer,host,null,{inWrapper:true,breakPages:true,ignoreWidth:false,ignoreHeight:false,useBase64URL:true});
    if(document.fonts?.ready)await document.fonts.ready;
    const sections=[...host.querySelectorAll('section.docx')];if(!sections.length)throw new Error('Word 文件中没有可转换的页面');
    const out=await PDFLib.PDFDocument.create();
    for(let i=0;i<sections.length;i++){$('pdfLabel').textContent=`正在转换 Word 第 ${i+1}/${sections.length} 页…`;const canvas=await window.html2canvas(sections[i],{scale:2,backgroundColor:'#ffffff',useCORS:true,logging:false}),png=await out.embedPng(new Uint8Array(await (await fetch(canvas.toDataURL('image/png'))).arrayBuffer())),pageWidth=595.28,pageHeight=pageWidth*canvas.height/canvas.width,page=out.addPage([pageWidth,pageHeight]);page.drawImage(png,{x:0,y:0,width:pageWidth,height:pageHeight});}
    out.setProducer('本地电子盖章工具 - Word 转换');return await out.save();
  }finally{host.remove();}
}

async function optimizeSignature(dataUrl){
  const img=await new Promise((res,rej)=>{const i=new Image();i.onload=()=>res(i);i.onerror=()=>rej(new Error('无法识别签名图片'));i.src=dataUrl}),max=1800,scale=Math.min(1,max/Math.max(img.naturalWidth,img.naturalHeight)),w=Math.max(1,Math.round(img.naturalWidth*scale)),h=Math.max(1,Math.round(img.naturalHeight*scale)),canvas=document.createElement('canvas');canvas.width=w;canvas.height=h;const ctx=canvas.getContext('2d',{willReadFrequently:true});ctx.drawImage(img,0,0,w,h);const data=ctx.getImageData(0,0,w,h),d=data.data;let minX=w,minY=h,maxX=-1,maxY=-1;
  for(let y=0;y<h;y++)for(let x=0;x<w;x++){const k=(y*w+x)*4,r=d[k],g=d[k+1],b=d[k+2],a=d[k+3],gray=(r+g+b)/3,darkness=255-gray,spread=Math.max(r,g,b)-Math.min(r,g,b),ink=Math.max(darkness,spread*.8),alpha=Math.max(0,Math.min(1,(ink-8)/65));d[k+3]=Math.round(a*Math.pow(alpha,.65));if(d[k+3]>12){minX=Math.min(minX,x);minY=Math.min(minY,y);maxX=Math.max(maxX,x);maxY=Math.max(maxY,y);}}
  if(maxX<minX)throw new Error('没有识别到手写笔迹');ctx.putImageData(data,0,0);const pad=Math.round(Math.max(4,(maxY-minY)*.06)),x0=Math.max(0,minX-pad),y0=Math.max(0,minY-pad),cw=Math.min(w-x0,maxX-minX+1+pad*2),ch=Math.min(h-y0,maxY-minY+1+pad*2),outW=Math.max(1200,cw),outH=Math.max(1,Math.round(outW*ch/cw)),out=document.createElement('canvas');out.width=outW;out.height=outH;const outCtx=out.getContext('2d');outCtx.imageSmoothingEnabled=true;outCtx.imageSmoothingQuality='high';outCtx.drawImage(canvas,x0,y0,cw,ch,0,0,outW,outH);return {dataUrl:out.toDataURL('image/png'),aspect:outW/outH};
}

async function renderPdf(){
  els.pages.innerHTML=''; els.empty.hidden=true;
  state.pdfDoc=await pdfjsLib.getDocument({data:state.pdfBytes.slice(),disableWorker:location.protocol==='file:'}).promise;
  $('pageCount').textContent=`${state.pdfDoc.numPages} 页`;
  for(let n=1;n<=state.pdfDoc.numPages;n++){
    const page=await state.pdfDoc.getPage(n), base=page.getViewport({scale:1}), scale=Math.min(1.45,900/base.width), viewport=page.getViewport({scale});
    const wrap=document.createElement('div'); wrap.className='page-wrap'; wrap.dataset.page=n;wrap.dataset.pdfWidth=base.width;wrap.dataset.pdfHeight=base.height; wrap.style.width=viewport.width+'px'; wrap.style.height=viewport.height+'px';
    const canvas=document.createElement('canvas'); const dpr=Math.min(devicePixelRatio||1,2); canvas.width=viewport.width*dpr;canvas.height=viewport.height*dpr;canvas.style.width=viewport.width+'px';canvas.style.height=viewport.height+'px';
    const number=document.createElement('span');number.className='page-number';number.textContent=`第 ${n} 页`;
    wrap.append(canvas,number);els.pages.append(wrap); await page.render({canvasContext:canvas.getContext('2d'),viewport,transform:dpr!==1?[dpr,0,0,dpr,0,0]:null}).promise;
    wrap.addEventListener('click',e=>{placeStamp(e,wrap,n);placeSignature(e,wrap,n);});
  }
  renderSeamPreview();
}

function friendlyError(err){const m=String(err?.message||err);if(/password/i.test(m))return 'PDF 有密码保护，请先解除密码';if(/worker|module|origin|fetch/i.test(m))return '请使用“启动盖章工具.command”打开，不要直接双击 index.html';return m.length>90?'文件损坏或格式暂不支持':m;}
function quantile(values,q){if(!values.length)return 0;values.sort((a,b)=>a-b);return values[Math.min(values.length-1,Math.max(0,Math.round((values.length-1)*q)))];}
function detectSealOuterCircle(data,width,height){
  const xs=[],ys=[],stride=Math.max(1,Math.floor(Math.sqrt(width*height/160000)));
  for(let y=0;y<height;y+=stride)for(let x=0;x<width;x+=stride){if(data[(y*width+x)*4+3]>14){xs.push(x);ys.push(y);}}
  if(xs.length<40)return null;
  const left=quantile(xs.slice(),.015),right=quantile(xs.slice(),.985),top=quantile(ys.slice(),.015),bottom=quantile(ys.slice(),.985),cx=(left+right)/2,cy=(top+bottom)/2,distances=[];
  for(let y=0;y<height;y+=stride)for(let x=0;x<width;x+=stride){if(data[(y*width+x)*4+3]>14)distances.push(Math.hypot(x-cx,y-cy));}
  const radius=quantile(distances,.99);
  if(!Number.isFinite(radius)||radius<10)return null;
  return {cx,cy,radius};
}
function refineSealAlpha(data,width,height,thinLevel,clarity){
  const count=width*height,src=new Uint8ClampedArray(count);for(let i=0;i<count;i++)src[i]=data[i*4+3];
  let alpha=src,radius=Math.min(8,Math.round(thinLevel*Math.max(width,height)/1100));
  if(radius>0){
    const horizontal=new Uint8ClampedArray(count),eroded=new Uint8ClampedArray(count);
    for(let y=0;y<height;y++)for(let x=0;x<width;x++){let v=255;for(let dx=-radius;dx<=radius;dx++){const xx=x+dx;if(xx<0||xx>=width){v=0;break;}v=Math.min(v,src[y*width+xx]);}horizontal[y*width+x]=v;}
    for(let y=0;y<height;y++)for(let x=0;x<width;x++){let v=255;for(let dy=-radius;dy<=radius;dy++){const yy=y+dy;if(yy<0||yy>=height){v=0;break;}v=Math.min(v,horizontal[yy*width+x]);}eroded[y*width+x]=v;}
    alpha=eroded;
  }
  const strength=clarity/100,contrast=1+strength*1.8;
  for(let i=0;i<count;i++){let v=alpha[i]/255;v=Math.max(0,Math.min(1,(v-.5)*contrast+.5));if(v>0)v=Math.pow(v,1-strength*.28);data[i*4+3]=Math.round(v*255);}
}
async function optimizeSeal(dataUrl){
  const img=await new Promise((res,rej)=>{const i=new Image();i.onload=()=>res(i);i.onerror=()=>rej(new Error('无法识别该图片'));i.src=dataUrl});
  const preserve=$('preserveSealOriginal').checked,max=preserve?4096:1800,scale=Math.min(1,max/Math.max(img.naturalWidth,img.naturalHeight)),w=Math.max(1,Math.round(img.naturalWidth*scale)),h=Math.max(1,Math.round(img.naturalHeight*scale));
  const source=document.createElement('canvas');source.width=w;source.height=h;const ctx=source.getContext('2d',{willReadFrequently:true});ctx.drawImage(img,0,0,w,h);
  const image=ctx.getImageData(0,0,w,h),d=image.data,remove=$('removeBackground').checked,extractRed=$('extractRedSeal').checked,t=+$('bgThreshold').value;
  let minX=w,minY=h,maxX=-1,maxY=-1;
  for(let y=0;y<h;y++)for(let x=0;x<w;x++){
    const k=(y*w+x)*4,r=d[k],g=d[k+1],b=d[k+2],originalAlpha=d[k+3],light=Math.min(r,g,b),spread=Math.max(r,g,b)-Math.min(r,g,b);
    if(extractRed){
      const dominance=r-Math.max(g,b),required=Math.max(5,30-t*.22),redFloor=Math.max(35,105-t*.55),soft=Math.max(8,28-t*.12);
      const colorScore=Math.min(1,Math.max(0,(dominance-required)/soft)),redScore=Math.min(1,Math.max(0,(r-redFloor)/35)),satScore=Math.min(1,spread/28),coverage=Math.pow(colorScore*redScore*satScore,preserve?.9:.48);
      d[k+3]=Math.round(originalAlpha*coverage);
      if(d[k+3]>0&&!preserve){const texture=.82+.18*(r/255);d[k]=Math.round(218*texture);d[k+1]=Math.round(30*texture);d[k+2]=Math.round(38*texture);}
    }else if(remove&&light>255-t&&spread<t*.55){
      d[k+3]=Math.round(originalAlpha*Math.max(0,(255-light)/Math.max(1,t)));
    }
  }
  refineSealAlpha(d,w,h,+$('sealThin').value,+$('sealClarity').value);
  minX=w;minY=h;maxX=-1;maxY=-1;for(let y=0;y<h;y++)for(let x=0;x<w;x++){if(d[(y*w+x)*4+3]>10){minX=Math.min(minX,x);minY=Math.min(minY,y);maxX=Math.max(maxX,x);maxY=Math.max(maxY,y);}}
  ctx.putImageData(image,0,0);if(maxX<minX){if(extractRed)throw new Error('没有识别到红色印章，请降低抠章强度或关闭智能抠取');return dataUrl;}const circle=extractRed?detectSealOuterCircle(d,w,h):null,fallbackSide=Math.max(maxX-minX+1,maxY-minY+1),side=Math.max(1,Math.ceil(circle?circle.radius*2:fallbackSide)),rawX=circle?circle.cx-circle.radius:minX-(side-(maxX-minX+1))/2,rawY=circle?circle.cy-circle.radius:minY-(side-(maxY-minY+1))/2,x0=preserve?Math.round(rawX):rawX,y0=preserve?Math.round(rawY):rawY,cropped=document.createElement('canvas');cropped.width=cropped.height=side;const cropCtx=cropped.getContext('2d');cropCtx.imageSmoothingEnabled=!preserve;cropCtx.drawImage(source,x0,y0,side,side,0,0,side,side);if(preserve)return cropped.toDataURL('image/png');const outputSize=Math.max(1200,Math.min(2400,side)),out=document.createElement('canvas');out.width=out.height=outputSize;const outCtx=out.getContext('2d');outCtx.imageSmoothingEnabled=true;outCtx.imageSmoothingQuality='high';outCtx.drawImage(cropped,0,0,outputSize,outputSize);return out.toDataURL('image/png');
}
function refreshSealImages(){document.querySelectorAll('.stamp img,.seam-preview img').forEach(i=>i.src=state.sealDataUrl);$('sealPreview').src=state.sealDataUrl;renderSeamPreview();}
$('reprocessBtn').onclick=async()=>{if(!state.originalSealDataUrl)return;state.sealDataUrl=await optimizeSeal(state.originalSealDataUrl);refreshSealImages();toast('已重新优化印章图片');};
$('bgThreshold').oninput=e=>$('thresholdValue').textContent=e.target.value;
$('sealClarity').oninput=e=>$('clarityValue').textContent=e.target.value+'%';
$('sealThin').oninput=e=>$('thinValue').textContent=e.target.value+' 级';
$('preserveSealOriginal').onchange=()=>toast($('preserveSealOriginal').checked?'已启用原图高清保留，请点击应用':'已切换为标准红色优化，请点击应用');
$('extractRedSeal').onchange=()=>{$('removeBackground').disabled=$('extractRedSeal').checked;};

function placeStamp(e,wrap,page){if(state.mode!=='normal'||!state.sealDataUrl||e.target.closest('.stamp'))return;const r=wrap.getBoundingClientRect(),random=$('randomTilt').checked?(Math.random()*4-2):0;const p={id:crypto.randomUUID(),page,cx:(e.clientX-r.left)/r.width,cy:(e.clientY-r.top)/r.height,sizeMm:+$('stampSize').value,rotation:(+$('stampRotation').value+random+360)%360,opacity:+$('stampOpacity').value/100};state.placements.push(p);selectStamp(p.id);drawStamp(wrap,p);}
function applyDepthStyle(el,depth){el.style.opacity=Math.min(1,depth);el.style.filter=depth>1?`brightness(${1-(depth-1)*.18}) contrast(${1+(depth-1)*.35}) saturate(${1+(depth-1)*.45})`:'';}
function layoutStamp(el,p,wrap){const pdfWidth=+wrap.dataset.pdfWidth,px=p.sizeMm*72/25.4*(wrap.clientWidth/pdfWidth);el.style.width=px+'px';el.style.height=px+'px';el.style.left=`calc(${p.cx*100}% - ${px/2}px)`;el.style.top=`calc(${p.cy*100}% - ${px/2}px)`;}
function drawStamp(wrap,p){const el=document.createElement('div');el.className='stamp';el.classList.toggle('ink-multiply',$('stampMultiply').checked);el.dataset.id=p.id;layoutStamp(el,p,wrap);el.style.transform=`rotate(${p.rotation||0}deg)`;applyDepthStyle(el,p.opacity??.9);el.classList.toggle('selected',state.selectedId===p.id);el.innerHTML=`<img src="${state.sealDataUrl}"><button title="删除">×</button>`;wrap.append(el);el.querySelector('button').onclick=()=>removeStamp(p.id);el.addEventListener('click',e=>{e.stopPropagation();selectStamp(p.id)});makeDraggable(el,p,wrap);}
function selectStamp(id){state.selectedId=id;document.querySelectorAll('.stamp').forEach(x=>x.classList.toggle('selected',x.dataset.id===id));const p=state.placements.find(x=>x.id===id);if(p){$('stampRotation').value=Math.round(p.rotation||0);$('rotationValue').textContent=(Math.round((p.rotation||0)*10)/10)+'°';$('stampOpacity').value=Math.round((p.opacity??.9)*100);$('opacityValue').textContent=Math.round((p.opacity??.9)*100)+'%';}}
function makeDraggable(el,p,wrap){el.addEventListener('pointerdown',e=>{if(e.target.tagName==='BUTTON')return;e.preventDefault();el.setPointerCapture(e.pointerId);el.classList.add('dragging');const r=wrap.getBoundingClientRect(),sx=e.clientX,sy=e.clientY,ocx=p.cx,ocy=p.cy;const move=ev=>{p.cx=Math.max(0,Math.min(1,ocx+(ev.clientX-sx)/r.width));p.cy=Math.max(0,Math.min(1,ocy+(ev.clientY-sy)/r.height));layoutStamp(el,p,wrap);};const up=()=>{el.classList.remove('dragging');el.removeEventListener('pointermove',move);el.removeEventListener('pointerup',up)};el.addEventListener('pointermove',move);el.addEventListener('pointerup',up);});}
function removeStamp(id){ state.placements=state.placements.filter(p=>p.id!==id);document.querySelector(`.stamp[data-id="${id}"]`)?.remove(); }
$('undoBtn').onclick=()=>{ const p=state.placements.at(-1);if(p)removeStamp(p.id); };

function setActiveSignature(slot){state.activeSignature=slot;const selected=state.signatureSlots[slot];state.signatureDataUrl=selected?.dataUrl||null;state.signatureAspect=selected?.aspect||2.5;$('signatureSlot1').classList.toggle('active',slot===0);$('signatureSlot2').classList.toggle('active',slot===1);}
$('signatureSlot1').onclick=()=>{setActiveSignature(0);toast(state.signatureDataUrl?'当前使用签名 1':'请先上传签名 1');};
$('signatureSlot2').onclick=()=>{setActiveSignature(1);toast(state.signatureDataUrl?'当前使用签名 2':'请先上传签名 2');};
function placeSignature(e,wrap,page){if(state.mode!=='signature'||e.target.closest('.signature-mark,.stamp'))return;if(!state.signatureDataUrl){toast(`请先上传签名 ${state.activeSignature+1}`);return;}const r=wrap.getBoundingClientRect(),p={id:crypto.randomUUID(),page,cx:(e.clientX-r.left)/r.width,cy:(e.clientY-r.top)/r.height,widthMm:+$('signatureSize').value,aspect:state.signatureAspect,opacity:+$('signatureOpacity').value/100,signatureSlot:state.activeSignature};state.signatures.push(p);drawSignature(wrap,p);selectSignature(p.id);}
function layoutSignature(el,p,wrap){const pdfWidth=+wrap.dataset.pdfWidth,widthPx=p.widthMm*72/25.4*(wrap.clientWidth/pdfWidth),heightPx=widthPx/p.aspect;el.style.width=widthPx+'px';el.style.height=heightPx+'px';el.style.left=`calc(${p.cx*100}% - ${widthPx/2}px)`;el.style.top=`calc(${p.cy*100}% - ${heightPx/2}px)`;el.style.opacity=p.opacity;}
function drawSignature(wrap,p){const slot=p.signatureSlot??0,source=state.signatureSlots[slot]?.dataUrl||state.signatureDataUrl;if(!source)return;const el=document.createElement('div');el.className='signature-mark';el.dataset.signatureId=p.id;el.dataset.signatureSlot=slot;layoutSignature(el,p,wrap);el.innerHTML=`<img src="${source}"><button title="删除签名">×</button>`;wrap.append(el);el.querySelector('button').onclick=()=>removeSignature(p.id);el.addEventListener('click',e=>{e.stopPropagation();selectSignature(p.id)});el.addEventListener('pointerdown',e=>{if(e.target.tagName==='BUTTON')return;e.preventDefault();el.setPointerCapture(e.pointerId);const r=wrap.getBoundingClientRect(),sx=e.clientX,sy=e.clientY,ocx=p.cx,ocy=p.cy;const move=ev=>{p.cx=Math.max(0,Math.min(1,ocx+(ev.clientX-sx)/r.width));p.cy=Math.max(0,Math.min(1,ocy+(ev.clientY-sy)/r.height));layoutSignature(el,p,wrap);};const up=()=>{el.removeEventListener('pointermove',move);el.removeEventListener('pointerup',up)};el.addEventListener('pointermove',move);el.addEventListener('pointerup',up);});}
function selectSignature(id){document.querySelectorAll('.signature-mark').forEach(x=>x.classList.toggle('selected',x.dataset.signatureId===id));}
function removeSignature(id){state.signatures=state.signatures.filter(p=>p.id!==id);document.querySelector(`.signature-mark[data-signature-id="${id}"]`)?.remove();}
$('undoSignatureBtn').onclick=()=>{const p=state.signatures.at(-1);if(p)removeSignature(p.id);};

function setMode(mode){state.mode=mode;$('normalMode').classList.toggle('active',mode==='normal');$('seamMode').classList.toggle('active',mode==='seam');$('signatureMode').classList.toggle('active',mode==='signature');$('normalControls').hidden=mode!=='normal';$('seamControls').hidden=mode!=='seam';$('signatureControls').hidden=mode!=='signature';document.querySelectorAll('.seam-preview').forEach(el=>el.classList.toggle('draggable',mode==='seam'));const messages={normal:'印章模式：点击页面插入印章',signature:'签名模式：点击页面插入手写签名',seam:'骑缝章模式：请直接拖动页面边缘的骑缝章'};toast(messages[mode]);}
$('normalMode').onclick=()=>setMode('normal');$('signatureMode').onclick=()=>setMode('signature');$('seamMode').onclick=()=>setMode('seam');
document.querySelectorAll('[data-edge]').forEach(b=>b.onclick=()=>{state.edge=b.dataset.edge;document.querySelectorAll('[data-edge]').forEach(x=>x.classList.toggle('active',x===b));renderSeamPreview();});
['seamOffset','seamOpacity','enableSeam'].forEach(id=>$(id).addEventListener('input',renderSeamPreview));
$('stampSize').oninput=e=>{const mm=+e.target.value;$('sizeValue').textContent=mm+' mm';$('seamSizeValue').textContent=mm+' mm（跟随全局）';state.placements.forEach(p=>{p.sizeMm=mm;const el=document.querySelector(`.stamp[data-id="${p.id}"]`),wrap=document.querySelector(`.page-wrap[data-page="${p.page}"]`);if(el&&wrap)layoutStamp(el,p,wrap);});renderSeamPreview();};
$('stampRotation').oninput=e=>{const value=+e.target.value;$('rotationValue').textContent=value+'°';const p=state.placements.find(x=>x.id===state.selectedId);if(p){p.rotation=value;const el=document.querySelector(`.stamp[data-id="${p.id}"]`);if(el)el.style.transform=`rotate(${value}deg)`;}};
$('stampOpacity').oninput=e=>{const value=+e.target.value;$('opacityValue').textContent=value+'%';const p=state.placements.find(x=>x.id===state.selectedId);if(p){p.opacity=value/100;const el=document.querySelector(`.stamp[data-id="${p.id}"]`);if(el)applyDepthStyle(el,p.opacity);}};
$('stampMultiply').onchange=e=>{document.querySelectorAll('.stamp,.seam-preview').forEach(el=>el.classList.toggle('ink-multiply',e.target.checked));toast(e.target.checked?'已启用真实叠色，黑字会从红章中透出':'已关闭真实叠色');};
$('seamOffset').oninput=e=>{$('seamOffsetValue').textContent=e.target.value+'%';renderSeamPreview();};
$('seamOpacity').oninput=e=>{$('seamOpacityValue').textContent=e.target.value+'%';renderSeamPreview();};
$('signatureSize').oninput=e=>{const mm=+e.target.value;$('signatureSizeValue').textContent=mm+' mm';state.signatures.forEach(p=>{p.widthMm=mm;const el=document.querySelector(`.signature-mark[data-signature-id="${p.id}"]`),wrap=document.querySelector(`.page-wrap[data-page="${p.page}"]`);if(el&&wrap)layoutSignature(el,p,wrap);});};
$('signatureOpacity').oninput=e=>{const opacity=+e.target.value/100;$('signatureOpacityValue').textContent=e.target.value+'%';state.signatures.forEach(p=>{p.opacity=opacity;const el=document.querySelector(`.signature-mark[data-signature-id="${p.id}"]`);if(el)el.style.opacity=opacity;});};

function updateSeamPreviewPosition(){document.querySelectorAll('.seam-preview').forEach(el=>{const wrap=el.parentElement,scale=wrap.clientWidth/(+wrap.dataset.pdfWidth),diameter=+$('stampSize').value*72/25.4*scale;el.style.left='';el.style.right='';el.style[state.edge]=(state.seamInsetMm*72/25.4*scale)+'px';el.style.top=`calc(${$('seamOffset').value}% - ${diameter/2}px)`;});}
function makeSeamDraggable(el,wrap){el.addEventListener('pointerdown',e=>{if(state.mode!=='seam')return;e.preventDefault();e.stopPropagation();el.setPointerCapture(e.pointerId);el.classList.add('dragging');const r=wrap.getBoundingClientRect(),startX=e.clientX,startY=e.clientY,startInset=state.seamInsetMm,startTop=+$('seamOffset').value,mmPerPx=(+wrap.dataset.pdfWidth)*25.4/72/r.width;const move=ev=>{const horizontal=(ev.clientX-startX)*mmPerPx*(state.edge==='right'?-1:1);state.seamInsetMm=Math.max(0,Math.min(60,startInset+horizontal));const top=Math.max(10,Math.min(85,startTop+(ev.clientY-startY)/r.height*100));$('seamOffset').value=top;$('seamOffsetValue').textContent=Math.round(top)+'%';$('seamPositionValue').textContent=state.seamInsetMm.toFixed(1)+' mm';updateSeamPreviewPosition();};const up=()=>{el.classList.remove('dragging');el.removeEventListener('pointermove',move);el.removeEventListener('pointerup',up);toast('骑缝章位置已记录，导出将使用此位置');};el.addEventListener('pointermove',move);el.addEventListener('pointerup',up);});}
function renderSeamPreview(){document.querySelectorAll('.seam-preview').forEach(e=>e.remove());if(!state.pdfDoc||!state.sealDataUrl||!$('enableSeam').checked)return;const count=state.pdfDoc.numPages,mm=+$('stampSize').value;document.querySelectorAll('.page-wrap').forEach((wrap,i)=>{const scale=wrap.clientWidth/(+wrap.dataset.pdfWidth),diameter=mm*72/25.4*scale,slice=diameter/count,inset=state.seamInsetMm*72/25.4*scale,el=document.createElement('div');el.className='seam-preview';el.classList.toggle('draggable',state.mode==='seam');el.classList.toggle('ink-multiply',$('stampMultiply').checked);el.style.width=slice+'px';el.style.height=diameter+'px';el.style.top=`calc(${$('seamOffset').value}% - ${diameter/2}px)`;el.style[state.edge]=inset+'px';applyDepthStyle(el,+$('seamOpacity').value/100);const img=document.createElement('img');img.src=state.sealDataUrl;img.style.width=diameter+'px';img.style.left=-(i*slice)+'px';el.append(img);wrap.append(el);makeSeamDraggable(el,wrap);});}

async function depthCanvas(dataUrl,depth=1){const img=await new Promise((res,rej)=>{const i=new Image();i.onload=()=>res(i);i.onerror=rej;i.src=dataUrl}),c=document.createElement('canvas');c.width=img.naturalWidth;c.height=img.naturalHeight;const ctx=c.getContext('2d',{willReadFrequently:true});ctx.drawImage(img,0,0);const data=ctx.getImageData(0,0,c.width,c.height),d=data.data,darken=Math.max(0,depth-1);for(let i=0;i<d.length;i+=4){d[i+3]=Math.min(255,d[i+3]*depth);if(darken){d[i]*=1-darken*.18;d[i+1]*=1-darken*.18;d[i+2]*=1-darken*.18;}}ctx.putImageData(data,0,0);return c;}
async function imageToPngBytes(dataUrl,depth=1){const c=await depthCanvas(dataUrl,depth);return new Uint8Array(await (await fetch(c.toDataURL('image/png'))).arrayBuffer());}
async function seamSlices(count,depth=1){ const source=await depthCanvas(state.sealDataUrl,depth),out=[];for(let i=0;i<count;i++){const x=Math.floor(i*source.width/count),next=Math.floor((i+1)*source.width/count),c=document.createElement('canvas');c.width=Math.max(1,next-x);c.height=source.height;c.getContext('2d').drawImage(source,x,0,next-x,source.height,0,0,c.width,c.height);out.push(new Uint8Array(await (await fetch(c.toDataURL('image/png'))).arrayBuffer()));}return out; }
function download(bytes,name,type){const blob=new Blob([bytes],{type}),url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);}

async function rasterizePdf(pdfBytes){
  const source=await pdfjsLib.getDocument({data:pdfBytes.slice()}).promise,out=await PDFLib.PDFDocument.create(),tiltMax=+$('pageTilt').value,blur=+$('pageBlur').value,renderScale=2.5,seamEnabled=$('enableSeam').checked,seamDepth=+$('seamOpacity').value/100,seamSource=seamEnabled?await depthCanvas(state.sealDataUrl,Math.max(1,seamDepth)):null;
  for(let n=1;n<=source.numPages;n++){
    els.export.textContent=`正在生成图片页 ${n}/${source.numPages}…`;
    const srcPage=await source.getPage(n),base=srcPage.getViewport({scale:1}),view=srcPage.getViewport({scale:renderScale}),raw=document.createElement('canvas');raw.width=Math.ceil(view.width);raw.height=Math.ceil(view.height);
    const rawCtx=raw.getContext('2d');await srcPage.render({canvasContext:rawCtx,viewport:view}).promise;
    if(seamSource){const count=source.numPages,pageIndex=n-1,sx=Math.floor(pageIndex*seamSource.width/count),sxNext=Math.floor(n*seamSource.width/count),sourceWidth=Math.max(1,sxNext-sx),diameter=+$('stampSize').value*72/25.4*renderScale,sliceWidth=Math.max(2,diameter/count),inset=state.seamInsetMm*72/25.4*renderScale,dx=state.edge==='right'?raw.width-sliceWidth-inset:inset,dy=raw.height*(+$('seamOffset').value/100)-diameter/2;rawCtx.save();rawCtx.globalAlpha=Math.min(1,seamDepth);rawCtx.globalCompositeOperation=$('stampMultiply').checked?'multiply':'source-over';rawCtx.drawImage(seamSource,sx,0,sourceWidth,seamSource.height,dx,dy,sliceWidth,diameter);rawCtx.restore();}
    const finalCanvas=document.createElement('canvas');finalCanvas.width=raw.width;finalCanvas.height=raw.height;const ctx=finalCanvas.getContext('2d');ctx.fillStyle='#fff';ctx.fillRect(0,0,finalCanvas.width,finalCanvas.height);const direction=Math.random()<.5?-1:1,angle=direction*Math.random()*tiltMax*Math.PI/180;
    ctx.save();ctx.translate(finalCanvas.width/2,finalCanvas.height/2);ctx.rotate(angle);ctx.scale(1.018,1.018);ctx.filter=blur>0?`blur(${blur*renderScale}px)`:'none';ctx.drawImage(raw,-raw.width/2,-raw.height/2);ctx.restore();
    const jpgBytes=new Uint8Array(await (await fetch(finalCanvas.toDataURL('image/jpeg',.94))).arrayBuffer()),jpg=await out.embedJpg(jpgBytes),target=out.addPage([base.width,base.height]);target.drawImage(jpg,{x:0,y:0,width:base.width,height:base.height});
  }
  out.setProducer('本地电子盖章工具 - 图片型 PDF');out.setModificationDate(new Date());return await out.save();
}

$('imagePdfMode').onchange=e=>$('scanEffectControls').hidden=!e.target.checked;
$('pageTilt').oninput=e=>$('pageTiltValue').textContent=e.target.value+'°';
$('pageBlur').oninput=e=>$('pageBlurValue').textContent=e.target.value+' px';

els.export.onclick=async()=>{try{els.export.disabled=true;els.export.textContent='正在生成…';const imageMode=$('imagePdfMode').checked,pdf=await PDFLib.PDFDocument.load(state.pdfBytes.slice()),blendMode=$('stampMultiply').checked?PDFLib.BlendMode.Multiply:PDFLib.BlendMode.Normal;for(const p of state.placements){const depth=p.opacity??.9,seal=await pdf.embedPng(await imageToPngBytes(state.sealDataUrl,depth)),page=pdf.getPage(p.page-1),{width,height}=page.getSize(),w=p.sizeMm*72/25.4,h=w,exportRotation=-(p.rotation||0),angle=exportRotation*Math.PI/180,cx=p.cx*width,cy=height-p.cy*height,x=cx-w/2*Math.cos(angle)+h/2*Math.sin(angle),y=cy-w/2*Math.sin(angle)-h/2*Math.cos(angle);page.drawImage(seal,{x,y,width:w,height:h,rotate:PDFLib.degrees(exportRotation),opacity:Math.min(1,depth),blendMode});}if(state.signatures.length){const embeddedSignatures={};for(const p of state.signatures){const slot=p.signatureSlot??0,source=state.signatureSlots[slot]?.dataUrl;if(!source)continue;if(!embeddedSignatures[slot])embeddedSignatures[slot]=await pdf.embedPng(await imageToPngBytes(source));const signature=embeddedSignatures[slot],page=pdf.getPage(p.page-1),{width,height}=page.getSize(),w=p.widthMm*72/25.4,h=w/p.aspect;page.drawImage(signature,{x:p.cx*width-w/2,y:height-p.cy*height-h/2,width:w,height:h,opacity:p.opacity});}}if($('enableSeam').checked&&!imageMode){const seamDepth=+$('seamOpacity').value/100,slices=await seamSlices(pdf.getPageCount(),seamDepth),diameterMm=+$('stampSize').value;for(let i=0;i<pdf.getPageCount();i++){const page=pdf.getPage(i),{x:pageX,y:pageY,width,height}=page.getCropBox(),h=diameterMm*72/25.4,w=h/pdf.getPageCount(),inset=state.seamInsetMm*72/25.4,img=await pdf.embedPng(slices[i]),y=pageY+height*(1-$('seamOffset').value/100)-h/2,x=state.edge==='right'?pageX+width-w-inset:pageX+inset;page.drawImage(img,{x,y,width:w,height:h,opacity:Math.min(1,seamDepth),blendMode});}}pdf.setProducer('本地电子盖章工具');pdf.setModificationDate(new Date());let bytes=await pdf.save();if(imageMode)bytes=await rasterizePdf(bytes);download(bytes,state.fileName.replace(/\.(pdf|docx)$/i,'')+(imageMode?'_图片版已盖章.pdf':'_已盖章.pdf'),'application/pdf');toast(imageMode?'图片型盖章 PDF 已生成':'盖章文件已生成');}catch(err){console.error(err);toast('导出失败：'+friendlyError(err));}finally{els.export.textContent='导出盖章文件';ready();}};

els.manifest.onclick=()=>{const data={document:state.fileName,pageCount:state.pdfDoc?.numPages||0,startedAt:state.startedAt,exportedAt:new Date().toISOString(),globalStampDiameterMm:+$('stampSize').value,normalStamps:state.placements.map(({id,...p})=>p),signatures:state.signatures.map(({id,...p})=>p),seamStamp:$('enableSeam').checked?{enabled:true,edge:state.edge,topPercent:+$('seamOffset').value,insetMm:state.seamInsetMm,diameterMm:+$('stampSize').value}:{enabled:false},imagePdf:$('imagePdfMode').checked?{enabled:true,renderDpi:180,maxRandomTiltDegrees:+$('pageTilt').value,blurPixels:+$('pageBlur').value}:{enabled:false},authorizationConfirmed:els.auth.checked,note:'本记录不包含文件、印章或签名图像内容。'};download(JSON.stringify(data,null,2),state.fileName.replace(/\.(pdf|docx)$/i,'')+'_盖章记录.json','application/json');};
ready();
$('removeBackground').disabled=$('extractRedSeal').checked;
