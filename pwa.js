let deferredInstallPrompt=null;
const installBtn=document.getElementById('installAppBtn');
function isStandalone(){return window.matchMedia('(display-mode: standalone)').matches||window.navigator.standalone===true}
function updateInstallButton(){
  if(!installBtn)return;
  if(isStandalone()){installBtn.style.display='none';return}
  installBtn.style.display='inline-flex';
}
window.addEventListener('beforeinstallprompt',e=>{
  e.preventDefault();
  deferredInstallPrompt=e;
  updateInstallButton();
});
window.addEventListener('appinstalled',()=>{
  deferredInstallPrompt=null;
  updateInstallButton();
});
installBtn?.addEventListener('click',async()=>{
  if(isStandalone())return;
  if(deferredInstallPrompt){
    deferredInstallPrompt.prompt();
    await deferredInstallPrompt.userChoice;
    deferredInstallPrompt=null;
    updateInstallButton();
    return;
  }
  const ua=navigator.userAgent||'';
  if(/iPhone|iPad|iPod/i.test(ua)){
    alert('iPhone/iPad: Safari의 공유 버튼 → 홈 화면에 추가를 선택하세요.');
  }else{
    alert('브라우저 메뉴에서 “앱 설치” 또는 “홈 화면에 추가”를 선택하세요.');
  }
});
if('serviceWorker' in navigator){
  window.addEventListener('load',async()=>{
    try{
      const reg=await navigator.serviceWorker.register('./sw.js?v=20261026-1',{scope:'./'});
      reg.update();
      if(reg.waiting)reg.waiting.postMessage('SKIP_WAITING');
      reg.addEventListener('updatefound',()=>{
        const worker=reg.installing;
        worker?.addEventListener('statechange',()=>{
          if(worker.state==='installed'&&navigator.serviceWorker.controller)worker.postMessage('SKIP_WAITING');
        });
      });
    }catch(err){console.warn('PWA service worker registration failed',err)}
  });
}
updateInstallButton();