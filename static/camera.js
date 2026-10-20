const v=document.getElementById('v'),out=document.getElementById('out'),canvas=document.createElement('canvas');
function snap(){canvas.width=v.videoWidth;canvas.height=v.videoHeight;canvas.getContext('2d').drawImage(v,0,0);return canvas.toDataURL('image/jpeg',0.85)}
async function post(url,body){const r=await fetch(url,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});return r.json()}
function say(t,c){out.textContent=t;out.className=c||''}
function startCam(cb){navigator.mediaDevices.getUserMedia({video:true}).then(s=>{v.srcObject=s;cb()}).catch(()=>say('Camera blocked or unavailable (use localhost or HTTPS).','unknown'))}
