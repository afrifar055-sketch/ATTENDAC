const v=document.getElementById('v'),out=document.getElementById('out'),canvas=document.createElement('canvas');
function snap(){canvas.width=v.videoWidth;canvas.height=v.videoHeight;canvas.getContext('2d').drawImage(v,0,0);return canvas.toDataURL('image/jpeg',0.85)}
async function post(url,body){const r=await fetch(url,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});return r.json()}
function say(t,c){out.textContent=t;out.className=c||''}
function startCam(cb){
 if(!navigator.mediaDevices){say('Camera needs a secure address. Open the site with https:// (or localhost).','unknown');return}
 navigator.mediaDevices.getUserMedia({video:true}).then(s=>{v.srcObject=s;cb()}).catch(e=>{
  const hint={NotAllowedError:'Camera permission is blocked. Click the lock icon next to the address, set Camera to Allow, then reload.',
   NotFoundError:'No camera was found on this device.',
   NotReadableError:'Another app or browser tab is using the camera. Close it, then reload.'}[e.name]||'Try reloading the page.';
  say('Camera error ('+e.name+'). '+hint,'unknown')})}
