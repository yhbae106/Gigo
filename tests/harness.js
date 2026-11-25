const fs=require('fs');
function ok(cond,msg){if(!cond){console.error('FAIL:',msg);process.exitCode=1}else console.log('OK:',msg)}
const html=fs.readFileSync('index.html','utf8');
const js=fs.readFileSync('app.js','utf8');
const cfg=fs.readFileSync('sync-config.js','utf8');
const manifest=JSON.parse(fs.readFileSync('manifest.webmanifest','utf8'));
const sw=fs.readFileSync('sw.js','utf8');
const pwa=fs.readFileSync('pwa.js','utf8');
[
  'filesInput','targetMonth','analyzeBtn','downloadBtn','previewCanvas','previewImage',
  'previewShell','previewEmpty','previewSub','matchInfo','saveJpgBtn',
  'masterPassword','masterLoginBtn','masterRefreshBtn','masterSaveBtn','masterStatus'
].concat(['sharedDownloadBtn','installAppBtn']).forEach(id=>ok(html.includes('id="'+id+'"'),'DOM id '+id));
ok(/app\.js\?v=/.test(html),'app.js cache-busting version');
ok(js.includes('async function renderExcelPreview'),'preview renderer exists');
ok(js.includes("toDataURL('image/jpeg',0.96)"),'JPG export encoding');
ok(js.includes("saveJpgBtn"),'JPG button wiring');
ok(js.includes('previewMerges')&&js.includes('previewMergeAt'),'merged-cell rendering');
ok(js.includes('gigoLoad')&&js.includes('gigoSave')&&js.includes('gigoMasterLogin'),'cross-browser master sync');
ok(js.includes('downloadSharedWorkbook')&&js.includes('sharedDownloadBtn'),'public shared workbook download');
ok(js.includes('serializeSharedSheet')&&js.includes('addSharedSourceSheet'),'shared workbook carries DCM/SAP source sheets');
ok(js.includes("source:{dcm:serializeSharedSheet(analysis.dws),sap:serializeSharedSheet(analysis.sws)}"),'master save stores DCM/SAP sheet payloads');
ok(/endpoint:\s*'https:\/\/script\.google\.com\/macros\/s\//.test(cfg),'Apps Script endpoint configured');
if(process.exitCode)process.exit(process.exitCode);

ok(html.includes('rel="manifest"'),'PWA manifest linked');
ok(html.includes('pwa.js?v='),'PWA installer script linked');
ok(manifest.display==='standalone','manifest standalone display');
ok(manifest.start_url==='./'&&manifest.scope==='./','manifest scoped for GitHub Pages path');
ok(Array.isArray(manifest.icons)&&manifest.icons.length>0,'manifest icon configured');
ok(sw.includes("self.addEventListener('install'")&&sw.includes("self.addEventListener('fetch'"),'service worker install/fetch handlers');
ok(pwa.includes('beforeinstallprompt')&&pwa.includes('serviceWorker.register'),'install prompt and service worker registration');
require('./preview-month')(js).then(()=>console.log('Gigo preview month regression passed.')).catch(error=>{
  console.error(error);
  process.exitCode=1;
});

require('./shell-cache')(sw,html).then(()=>console.log('Gigo cache scope and offline assets passed.')).catch(error=>{
  console.error(error);
  process.exitCode=1;
});
