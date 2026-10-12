/* 同色系识别与文字遮挡重建增强层 */
function smartHsv(r,g,b){r/=255;g/=255;b/=255;const hi=Math.max(r,g,b),lo=Math.min(r,g,b),d=hi-lo;let h=0;if(d){if(hi===r)h=((g-b)/d)%6;else if(hi===g)h=(b-r)/d+2;else h=(r-g)/d+4;h=(h*60+360)%360}return [h,hi?d/hi:0,hi]}
function hueGap(a,b){const d=Math.abs(a-b);return Math.min(d,360-d)}
function fillOccludedRuns(mask,w,h){
  const out=mask.slice(),maxX=Math.max(10,Math.round(w*.28)),maxY=Math.max(10,Math.round(h*.20));
  for(let y=0;y<h;y++){
    let previous=-1;
    for(let x=0;x<w;x++)if(mask[y*w+x]){
      if(previous>=0&&x-previous-1<=maxX)for(let q=previous+1;q<x;q++)out[y*w+q]=1;
      previous=x;
    }
  }
  for(let x=0;x<w;x++){
    let previous=-1;
    for(let y=0;y<h;y++)if(mask[y*w+x]){
      if(previous>=0&&y-previous-1<=maxY)for(let q=previous+1;q<y;q++)out[q*w+x]=1;
      previous=y;
    }
  }
  return out;
}
function smartProcess(push=false){
  if(!state.loaded)return;
  const w=src.width,h=src.height,img=s.getImageData(0,0,w,h),tol=+$('tolerance').value;
  const colors=state.include.length?state.include:[hexRgb($('colorInput').value)];
  const labs=colors.map(c=>rgbLab(...c)),hsvs=colors.map(c=>smartHsv(...c));
  const rejectLabs=state.exclude.map(c=>rgbLab(...c)),rejectHsvs=state.exclude.map(c=>smartHsv(...c));
  const family=$('hueFamily').checked,raw=new Uint8Array(w*h),rawAlpha=[];
  for(let i=0;i<raw.length;i++){
    const k=i*4,r=img.data[k],g=img.data[k+1],b=img.data[k+2],a=img.data[k+3],lab=rgbLab(r,g,b),hsv=smartHsv(r,g,b);
    const perceptual=labs.some(v=>deltaE(lab,v)<=tol);
    const sameFamily=family&&hsv[1]>.10&&hsvs.some(v=>hueGap(hsv[0],v[0])<=Math.max(5,tol*.72)&&Math.abs(hsv[1]-v[1])<.72);
    const rejected=rejectLabs.some(v=>deltaE(lab,v)<=tol)||rejectHsvs.some(v=>hsv[1]>.08&&hueGap(hsv[0],v[0])<Math.max(4,tol*.55));
    const colorfulTarget=hsvs.some(v=>v[1]>.25),blackInk=colorfulTarget&&(hsv[2]<.24||(hsv[1]<.07&&hsv[2]<.58));
    raw[i]=+((perceptual||sameFamily)&&!rejected&&!blackInk);
    if(raw[i])rawAlpha.push(a);
  }
  state.targetAlpha=rawAlpha.length?rawAlpha.sort((a,b)=>a-b)[rawAlpha.length>>1]:255;
  let m=raw.slice(),bridge=+$('bridge').value,smooth=+$('smooth').value;
  if(bridge)m=integralMorph(integralMorph(m,w,h,bridge,true),w,h,bridge,false);
  if(smooth)m=integralMorph(integralMorph(m,w,h,smooth,true),w,h,smooth,false);
  removeSpecks(m,w,h,+$('despeckle').value);
  if($('fillHoles').checked){m=fillOccludedRuns(m,w,h);fillHoles(m,w,h,Math.max(300,w*h*.10));}
  if($('invert').checked)for(let i=0;i<m.length;i++)m[i]=1-m[i];
  state.rawMask=raw;state.mask=m;state.synthetic=new Uint8Array(m.length);
  for(let i=0;i<m.length;i++)state.synthetic[i]=+(m[i]&&!raw[i]);
  state.inferred=state.completion?inferGap(m,w,h):new Uint8Array(m.length);
  smartRender();if(push)snapshot();
  const real=raw.reduce((a,b)=>a+b,0),rebuilt=state.synthetic.reduce((a,b)=>a+b,0);
  $('processStatus').textContent=`同色系 ${real.toLocaleString()} 像素 · 黑字/孔洞补全 ${rebuilt.toLocaleString()} 像素`;
  $('status').textContent='抠图完成，可下载透明 PNG';
}
function smartRender(){
  if(!state.mask)return;
  const w=src.width,h=src.height,base=s.getImageData(0,0,w,h),im=new ImageData(w,h),picked=hexRgb($('colorInput').value),feather=+$('feather').value;
  const expanded=feather?integralMorph(state.mask,w,h,feather,true):state.mask;
  for(let i=0;i<state.mask.length;i++){
    const k=i*4,real=state.rawMask?.[i]&&state.mask[i],rebuilt=state.synthetic?.[i],guess=state.inferred?.[i];
    if(real){im.data[k]=base.data[k];im.data[k+1]=base.data[k+1];im.data[k+2]=base.data[k+2];im.data[k+3]=base.data[k+3]}
    else if(rebuilt||guess){im.data[k]=picked[0];im.data[k+1]=picked[1];im.data[k+2]=picked[2];im.data[k+3]=state.targetAlpha||255}
    else if(expanded[i]){im.data[k]=picked[0];im.data[k+1]=picked[1];im.data[k+2]=picked[2];im.data[k+3]=72}
  }
  o.putImageData(state.show==='original'?base:im,0,0);x.clearRect(0,0,w,h);
  if(state.show!=='original'&&state.completion){x.fillStyle='rgba(255,181,36,.36)';for(let i=0;i<state.inferred.length;i++)if(state.inferred[i])x.fillRect(i%w,i/w|0,1,1);x.fillStyle='#1d2521';x.font=`bold ${Math.max(12,w/38)}px sans-serif`;x.fillText('AI推测内容',14,30)}
}
process=smartProcess;render=smartRender;
$('hueFamily').onchange=()=>process(true);
if(state.loaded)process(true);
