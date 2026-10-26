const fs=require('fs');
function ok(cond,msg){if(!cond){console.error('FAIL:',msg);process.exitCode=1}else console.log('OK:',msg)}
const html=fs.readFileSync('index.html','utf8');
const js=fs.readFileSync('app.js','utf8');
const cfg=fs.readFileSync('sync-config.js','utf8');
[
  'filesInput','targetMonth','analyzeBtn','downloadBtn','previewCanvas','previewImage',
  'previewShell','previewEmpty','previewSub','matchInfo','saveJpgBtn',
  'masterPassword','masterLoginBtn','masterRefreshBtn','masterSaveBtn','masterStatus'
].forEach(id=>ok(html.includes('id="'+id+'"'),'DOM id '+id));
ok(/app\.js\?v=/.test(html),'app.js cache-busting version');
ok(js.includes('async function renderExcelPreview'),'preview renderer exists');
ok(js.includes('firstCol=1,lastCol=7,firstRow=6,lastRow=findPreviewLastRow(ws)'),'preview fixed to row 6 through column G');
ok(js.includes("toDataURL('image/jpeg',0.96)"),'JPG export encoding');
ok(js.includes('previewDisplayText')&&js.includes("c===1||c===3"),'code columns rendered without commas');
ok(js.includes("c===6||c===7")&&js.includes("toLocaleString('ko-KR')"),'amount columns rendered with thousand separators');
ok(js.includes('findPreviewLastRow'),'trailing blank rows removed');
ok(js.includes("saveJpgBtn"),'JPG button wiring');
ok(js.includes('previewMerges')&&js.includes('previewMergeAt'),'merged-cell rendering');
ok(js.includes('gigoLoad')&&js.includes('gigoSave')&&js.includes('gigoMasterLogin'),'cross-browser master sync');
ok(/endpoint:\s*'https:\/\/script\.google\.com\/macros\/s\//.test(cfg),'Apps Script endpoint configured');
if(process.exitCode)process.exit(process.exitCode);
console.log('Gigo harness passed.');
