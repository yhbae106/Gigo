
const $=id=>document.getElementById(id);
let dcmFile=null,sapFile=null,analysis=null,previewJpgUrl='';
let gigoShared={version:1,months:{}}, masterOn=false;
const T=[
['약국현금','현금','10039905','대구가톨릭대학교의료원','DW35'],['약국현금','현금','10047248','부산성모병원(재단법인천주교부산교구유지재단)','DW35'],['약국현금','현금','10039905','대구가톨릭대학교의료원','DW30'],['약국현금','현금','10040464','(재)천주교부산교구메리놀병원','DW30'],['약국현금','현금','10040911','창원파티마병원','DW30'],['약국현금','현금','10047248','부산성모병원(재단법인천주교부산교구유지재단)','DW30'],['약국현금','현금','10093445','주식회사기고(GIGO)','DW18'],
['병원현금','현금','10039568','포항성모병원','DW23'],['병원현금','현금','10039568','포항성모병원','DW26'],['병원현금','현금','10040464','(재)천주교부산교구메리놀병원','DW23'],['병원현금','현금','10039205','대구파티마병원','DW34'],['병원현금','현금','10039205','대구파티마병원','DW24'],['병원현금','현금','10039425','대구가톨릭대학교칠곡가톨릭병원','DW24'],['병원현금','현금','10039905','대구가톨릭대학교의료원','DW24'],['병원현금','현금','10040464','(재)천주교부산교구메리놀병원','DW24'],['병원현금','현금','10040911','창원파티마병원','DW24'],['병원현금','현금','10047248','부산성모병원(재단법인천주교부산교구유지재단)','DW24'],['병원현금','현금','10093445','주식회사기고(GIGO)','DW24'],['병원현금','현금','10040464','(재)천주교부산교구메리놀병원','DW26'],['병원현금','현금','10040911','창원파티마병원','DW26'],['병원현금','현금','10047248','부산성모병원(재단법인천주교부산교구유지재단)','DW26'],
['도입현금','현금','10093445','주식회사기고(GIGO)','DW21'],['도입현금','현금','10093445','주식회사기고','DW22'],['도입현금','현금','10039205','대구파티마병원','DW28'],['도입현금','현금','10039568','포항성모병원','DW28'],['도입현금','현금','10039905','대구가톨릭대학교의료원','DW28'],['도입현금','현금','10040464','(재)천주교부산교구메리놀병원','DW28'],['도입현금','현금','10040911','창원파티마병원','DW28'],['도입현금','현금','10047248','부산성모병원(재단법인천주교부산교구유지재단)','DW28'],
['도입어음','어음','10039205','대구파티마병원','DW28'],['도입어음','어음','10039568','포항성모병원','DW28'],['도입어음','어음','10039905','대구가톨릭대학교의료원','DW28'],['도입어음','어음','10040464','(재)천주교부산교구메리놀병원','DW28'],['도입어음','어음','10040911','창원파티마병원','DW28'],['도입어음','어음','10047248','부산성모병원(재단법인천주교부산교구유지재단)','DW28']
].map((x,i)=>({id:i+1,group:x[0],payType:x[1],code:x[2],name:x[3],ch:x[4]}));
const G=['약국현금','병원현금','도입현금','도입어음'], EX=new Set(['10040464','10047248']), SCH=new Set(['DW21','DW22','DW28']);
const code=v=>String(v==null?'':v).replace(/[^0-9]/g,'').replace(/^0+/,'')||'';
const num=v=>{const n=Number(String(v==null?'':v).replace(/,/g,''));return Number.isFinite(n)?n:0};
const fmt=v=>Math.round(v).toLocaleString('ko-KR');
const val=(ws,r,c)=>{const x=ws.getCell(r,c).value;if(x&&typeof x==='object'&&'result'in x)return x.result;if(x&&typeof x==='object'&&'text'in x)return x.text;if(x&&typeof x==='object'&&Array.isArray(x.richText))return x.richText.map(t=>t.text).join('');return x};
function status(m,k){$('status').textContent=m;$('status').className='status'+(k?' '+k:'')}
function strictKind(wb){
  const d=wb.worksheets.find(s=>String(val(s,5,2)).includes('실납코드')&&String(val(s,5,4)).includes('채널'));
  const p=wb.worksheets.find(s=>String(val(s,1,17)).includes('실납처')&&String(val(s,1,22)).includes('채널'));
  if(d&&!p)return 'dcm';
  if(p&&!d)return 'sap';
  return null;
}
function setFileBox(kind,file){
  const box=$(kind+'Box'),name=$(kind+'Name');
  if(file){box.classList.remove('wait');box.classList.add('ok');name.textContent='✓ '+file.name}
  else{box.classList.remove('ok');box.classList.add('wait');name.textContent='인식 대기'}
}
async function acceptFiles(list){
  const files=Array.from(list||[]).filter(f=>/\.xlsx$/i.test(f.name));
  if(files.length!==2){status('DCM 수금내역서와 SAP매출 .xlsx 파일 2개를 함께 선택해 주세요.','bad');return}
  dcmFile=null;sapFile=null;analysis=null;previewJpgUrl='';if($('saveJpgBtn'))$('saveJpgBtn').disabled=true;setFileBox('dcm',null);setFileBox('sap',null);
  $('downloadBtn').disabled=true;$('analyzeBtn').disabled=true;
  status('두 파일의 구조를 확인해 DCM과 SAP을 자동 인식하고 있습니다.');
  try{
    const loaded=await Promise.all(files.map(async file=>({file,wb:await load(file)})));
    for(const x of loaded){
      const kind=strictKind(x.wb);
      if(kind==='dcm'&&!dcmFile)dcmFile=x.file;
      else if(kind==='sap'&&!sapFile)sapFile=x.file;
    }
    if(!dcmFile||!sapFile)throw new Error('두 파일 중 DCM 또는 SAP 형식을 인식하지 못했습니다.');
    setFileBox('dcm',dcmFile);setFileBox('sap',sapFile);tryAutoMonthFromNames();
    $('analyzeBtn').disabled=false;
    await analyzeFiles();
  }catch(e){status('파일 인식 오류: '+e.message,'bad')}
}
const filesInput=$('filesInput'),filesDrop=$('filesDrop');
filesInput.addEventListener('change',e=>acceptFiles(e.target.files));
['dragenter','dragover'].forEach(ev=>filesDrop.addEventListener(ev,e=>{e.preventDefault();filesDrop.classList.add('drag')}));
['dragleave','drop'].forEach(ev=>filesDrop.addEventListener(ev,e=>{e.preventDefault();filesDrop.classList.remove('drag')}));
filesDrop.addEventListener('drop',e=>acceptFiles(e.dataTransfer.files));

function syncEndpoint(){return String(window.GIGO_SYNC?.endpoint||'').trim()}
function masterStatus(msg,kind=''){const el=$('masterStatus');if(!el)return;el.textContent=msg;el.className='status'+(kind?' '+kind:'')}
function normalizeShared(v){
  const out={version:1,months:{}};
  if(v&&typeof v==='object'&&v.months&&typeof v.months==='object'){
    for(const [m,x] of Object.entries(v.months)){
      if(!/^20\d{2}-\d{2}$/.test(m)||!x||!Array.isArray(x.rows))continue;
      out.months[m]={rows:x.rows.map(r=>({...r,amount:num(r.amount),final:num(r.final)})),matched:num(x.matched),unmatched:num(x.unmatched),savedAt:x.savedAt||''};
    }
  }
  return out;
}
async function syncRequest(payload){
  if(!syncEndpoint())throw new Error('공용 저장소 주소가 아직 설정되지 않았습니다.');
  const res=await fetch(syncEndpoint(),{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify(payload)});
  if(!res.ok)throw new Error('공용 저장소 HTTP '+res.status);
  const json=await res.json();
  if(json?.ok!==true)throw new Error(json?.error||'공용 저장소 요청 실패');
  return json;
}
function setMasterUI(on){
  masterOn=!!on;
  if($('masterBadge'))$('masterBadge').textContent=masterOn?'MASTER ON':'읽기 전용';
  if($('masterLoginBtn'))$('masterLoginBtn').textContent=masterOn?'마스터 종료':'마스터 로그인';
  if($('masterPassword'))$('masterPassword').disabled=masterOn;
  if($('masterSaveBtn'))$('masterSaveBtn').disabled=!(masterOn&&analysis);
}
async function pullSharedData(render=true){
  if(!syncEndpoint()){masterStatus('공용 저장소 미연결 · sync-config.js에 Apps Script URL 설정 필요');return false}
  try{
    masterStatus('공용 저장 데이터 확인 중…');
    const json=await syncRequest({type:'gigoLoad'});
    gigoShared=normalizeShared(json.data);
    const keys=Object.keys(gigoShared.months).sort();
    if(!keys.length){masterStatus('공용 저장 데이터가 아직 없습니다.');return true}
    const latest=keys.at(-1),snap=gigoShared.months[latest];
    masterStatus('공용 데이터 적용 · 최근 저장 '+latest+(json.updatedAt?' · '+new Date(json.updatedAt).toLocaleString('ko-KR'):''),'good');
    if(render&&!analysis){
      $('targetMonth').value=latest;
      const c={rows:snap.rows,matched:snap.matched||snap.rows.length,unmatched:snap.unmatched||0};
      show(c);
      await renderExcelPreview(c,latest);
      $('matchInfo').textContent='공용 저장 · '+latest;
    }
    return true;
  }catch(e){masterStatus('공용 데이터 조회 실패 · '+e.message,'bad');return false}
}
async function masterLogin(){
  if(masterOn){sessionStorage.removeItem('gigo-master-password');setMasterUI(false);masterStatus('마스터 모드를 종료했습니다.');return}
  const pw=String($('masterPassword')?.value||'').trim();
  if(!pw){masterStatus('마스터 비밀번호를 입력해 주세요.','bad');return}
  try{
    masterStatus('마스터 인증 중…');
    await syncRequest({type:'gigoMasterLogin',password:pw});
    sessionStorage.setItem('gigo-master-password',pw);
    setMasterUI(true);
    $('masterPassword').value='';
    masterStatus('마스터 인증 완료 · 현재 월 결과를 공용 저장할 수 있습니다.','good');
  }catch(e){sessionStorage.removeItem('gigo-master-password');setMasterUI(false);masterStatus('마스터 인증 실패 · '+e.message,'bad')}
}
async function saveMasterCurrent(){
  if(!masterOn||!analysis)return;
  const pw=sessionStorage.getItem('gigo-master-password')||'';
  const month=$('targetMonth').value||analysis.targetMonth||'';
  try{
    $('masterSaveBtn').disabled=true;masterStatus(month+' 공용 저장 중…');
    const snap={rows:analysis.calc.rows.map(x=>({id:x.id,group:x.group,payType:x.payType,code:x.code,name:x.name,ch:x.ch,amount:x.amount,final:x.final,source:x.source,found:x.found,isExc:x.isExc})),matched:analysis.calc.matched,unmatched:analysis.calc.unmatched};
    const json=await syncRequest({type:'gigoSave',password:pw,month,snapshot:snap});
    gigoShared=normalizeShared(json.data);
    masterStatus(month+' 공용 저장 완료 · 다른 브라우저에서 자동 조회됩니다.','good');
  }catch(e){masterStatus('공용 저장 실패 · '+e.message,'bad')}
  finally{$('masterSaveBtn').disabled=!(masterOn&&analysis)}
}
function parseYmText(text){
  const t=String(text||'');
  let m=t.match(/(20\d{2})[^0-9]{0,3}(1[0-2]|0?[1-9])\s*월?/);
  if(m)return {year:Number(m[1]),month:Number(m[2])};
  m=t.match(/(?:^|[^0-9])(\d{2})[._\-\s](1[0-2]|0?[1-9])(?:[^0-9]|$)/);
  if(m)return {year:2000+Number(m[1]),month:Number(m[2])};
  m=t.match(/(?:^|[^0-9])(1[0-2]|0?[1-9])\s*월/);
  if(m)return {year:new Date().getFullYear(),month:Number(m[1])};
  return null;
}
function addMonth(ym,delta){
  const d=new Date(ym.year,ym.month-1+delta,1);
  return {year:d.getFullYear(),month:d.getMonth()+1};
}
function ymString(ym){return ym.year+'-'+String(ym.month).padStart(2,'0')}
function workbookYm(file,wb){
  const texts=[file?.name||'',...wb.worksheets.map(x=>x.name||'')];
  for(const t of texts){const ym=parseYmText(t);if(ym)return ym}
  return null;
}
function tryAutoMonthFromNames(){
  const d=dcmFile?parseYmText(dcmFile.name):null;
  const s=sapFile?parseYmText(sapFile.name):null;
  let target=d||(s?addMonth(s,1):null);
  if(target)$('targetMonth').value=ymString(target);
}
const now=new Date();$('targetMonth').value=now.getFullYear()+'-'+String(now.getMonth()+1).padStart(2,'0');
async function load(f){const w=new ExcelJS.Workbook();await w.xlsx.load(await f.arrayBuffer());return w}
function sheet(w,k){if(k==='dcm')return w.worksheets.find(s=>String(val(s,5,2)).includes('실납코드')&&String(val(s,5,4)).includes('채널'))||w.worksheets[0];return w.worksheets.find(s=>String(val(s,1,17)).includes('실납처')&&String(val(s,1,22)).includes('채널'))||w.worksheets[0]}
function dmap(ws){const p=new Map(),t=new Map();let rows=0;for(let r=8;r<=ws.rowCount;r++){const c=code(val(ws,r,2)),h=String(val(ws,r,4)||'').trim();if(!c||!h)continue;rows++;const k=c+'|'+h;p.set(k,(p.get(k)||0)+num(val(ws,r,12)));t.set(k,(t.get(k)||0)+num(val(ws,r,14)))}return{prev:p,total:t,rows}}
function smap(ws){const cash=new Map(),note=new Map();let rows=0;for(let r=2;r<=ws.rowCount;r++){const c=code(val(ws,r,17)),h=String(val(ws,r,22)||'').trim();if(!c||!h)continue;const pr=String(val(ws,r,25)||''),cashFlag=pr.includes('콩코르')||pr.includes('크레스토'),amt=num(val(ws,r,49))+num(val(ws,r,52)),k=c+'|'+h,m=cashFlag?cash:note;rows++;m.set(k,(m.get(k)||0)+amt)}return{cash,note,rows}}
function calc(d,s){const rows=[];let matched=0,unmatched=0;for(const x of T){const k=x.code+'|'+x.ch,exc=(x.ch==='DW24'&&EX.has(x.code))||(x.code==='10093445'&&(x.ch==='DW21'||x.ch==='DW22'));let amt=0,src='';if(exc){amt=d.total.get(k)||0;src='DCM 합계'}else if(SCH.has(x.ch)){amt=(x.payType==='현금'?s.cash:s.note).get(k)||0;src='SAP '+x.payType}else{amt=d.prev.get(k)||0;src='DCM 전월'}const found=amt!==0||(exc?d.total.has(k):SCH.has(x.ch)?(x.payType==='현금'?s.cash:s.note).has(k):d.prev.has(k));found?matched++:unmatched++;rows.push({...x,amount:amt,final:amt,source:src,found,isExc:exc})}for(const g of G){const a=rows.filter(x=>x.group===g),sum=a.reduce((q,x)=>q+x.amount,0),diff=(sum<=0?0:Math.ceil(sum/1000000)*1000000)-sum;if(diff>0&&a.length){let m=a[0];for(const x of a)if(x.amount>m.amount)m=x;m.final+=diff}}return{rows,matched,unmatched}}
function show(a){
  $('summary').style.display='grid';$('resultCard').style.display='block';
  const z={};G.forEach(g=>z[g]=a.rows.filter(x=>x.group===g).reduce((q,x)=>q+x.final,0));
  const cash=z['약국현금']+z['병원현금']+z['도입현금'],note=z['도입어음'];
  $('mCash').textContent=fmt(cash);$('mNote').textContent=fmt(note);$('mTotal').textContent=fmt(cash+note);
  $('matchInfo').textContent='매칭 '+a.matched+'건 · 미매칭 '+a.unmatched+'건';
  $('resultBody').innerHTML=a.rows.map(x=>'<tr><td>'+x.group+'</td><td>'+x.name+'</td><td>'+x.ch+'</td><td class="num">'+fmt(x.amount)+'</td><td class="num">'+fmt(x.final)+'</td><td><span class="tag '+(!x.found?'none':x.isExc?'exc':x.source.startsWith('SAP')?'sap':'dcm')+'">'+(!x.found?'미매칭':x.source)+'</span></td></tr>').join('');
}
async function analyzeFiles(){
  try{
    if(!dcmFile||!sapFile)throw new Error('DCM과 SAP 파일 2개가 모두 필요합니다.');
    $('analyzeBtn').disabled=true;status('대상 월 판정 · 수금 계산 · 엑셀 미리보기를 생성하고 있습니다.');
    const [dw,sw]=await Promise.all([load(dcmFile),load(sapFile)]),ds=sheet(dw,'dcm'),ss=sheet(sw,'sap');
    if(!ds||!ss)throw new Error('필수 시트를 찾을 수 없습니다.');
    const dYm=workbookYm(dcmFile,dw),sYm=workbookYm(sapFile,sw);
    let target=dYm||(sYm?addMonth(sYm,1):null);
    if(target)$('targetMonth').value=ymString(target);
    let monthMsg='';
    if(dYm&&sYm){
      const expect=addMonth(sYm,1);
      if(dYm.year===expect.year&&dYm.month===expect.month)monthMsg='대상월 '+dYm.year+'년 '+dYm.month+'월 · DCM 당월 / SAP 전월 정상';
      else monthMsg='월 확인 필요 · DCM '+dYm.month+'월 / SAP '+sYm.month+'월';
    }else if(dYm)monthMsg='대상월 '+dYm.year+'년 '+dYm.month+'월 · DCM 기준';
    else if(sYm)monthMsg='대상월 '+target.year+'년 '+target.month+'월 · SAP 전월 기준';
    const d=dmap(ds),sm=smap(ss),c=calc(d,sm);
    analysis={calc:c,dws:ds,sws:ss,targetMonth:$('targetMonth').value};
    show(c);$('downloadBtn').disabled=false;if($('masterSaveBtn'))$('masterSaveBtn').disabled=!masterOn;
    await renderExcelPreview(c,$('targetMonth').value);
    status((c.unmatched?'계산 완료 · '+c.matched+'건 매칭 / '+c.unmatched+'건 미매칭':'계산 완료 · '+c.matched+'건 모두 매칭')+' · '+monthMsg,c.unmatched?'':'good');
  }catch(e){console.error(e);status('오류: '+e.message,'bad')}
  finally{$('analyzeBtn').disabled=!(dcmFile&&sapFile)}
}
$('analyzeBtn').addEventListener('click',analyzeFiles);

function copy(src,wb,name){
  const dst=wb.addWorksheet(name);
  for(let r=1;r<=src.rowCount;r++){
    const a=[];
    for(let c=1;c<=src.columnCount;c++) a.push(val(src,r,c));
    dst.addRow(a);
  }
  return dst;
}
function cloneStyle(s){return s?JSON.parse(JSON.stringify(s)):{}}
const BASE10=[
[-620,-620],[-9963130,-9963130],[1973867,1973867],[5013161,5013161],[8649953,8649953],[23216777,23427079],[-100310,-100310],
[1434756,1434756],[0,0],[0,0],[0,0],[1869038,1869038],[1025264,1025264],[2262722,2262722],[-2375,-2375],[3435165,3494260],[-533029,-533029],[3202654,3202654],[0,0],[1233804,1233804],[12906,12906],
[-95,-95],[0,0],[1244160,1244160],[3110400,3110400],[8234460,8234460],[0,0],[13213260,13411075],[0,0],
[7440560,7440560],[32687074,32687074],[3117408,3117408],[0,0],[39993296,40754958],[0,0]
];
const POS=[7,8,9,10,11,12,13,17,18,19,20,21,22,23,24,25,26,27,28,29,30,37,38,39,40,41,42,43,44,46,47,48,49,50,51];
const GROUP_ROWS={약국현금:16,병원현금:36,도입현금:45,도입어음:52};
const NF='_-* #,##0_-;\\-* #,##0_-;_-* "-"_-;_-@_-';
function borderAll(){
  const x={style:'thin',color:{argb:'FF000000'}};
  return {top:x,left:x,bottom:x,right:x};
}
function setBaseRow(ws,r,vals){
  vals.forEach((v,i)=>ws.getCell(r,i+1).value=v);
}
function buildBaseWorkbook(){
  const wb=new ExcelJS.Workbook();
  wb.creator='Gigo Collection Dashboard';
  const ws=wb.addWorksheet('수금내역서');
  const widths=[12.19921875,10.59765625,10.19921875,36.796875,9.59765625,14,17.09765625,14,8.796875,8.796875,13,13,13,13];
  widths.forEach((w,i)=>ws.getColumn(i+1).width=w);
  for(let r=1;r<=52;r++) ws.getRow(r).height=22.8;
  ws.getRow(53).height=16.2; ws.getRow(54).height=16.2;
  [14,15,31,32,33,34,35].forEach(r=>ws.getRow(r).hidden=true);
  ws.mergeCells('D1:E3');
  ws.mergeCells('B16:E16'); ws.mergeCells('B36:E36'); ws.mergeCells('B45:E45'); ws.mergeCells('B52:E52'); ws.mergeCells('D54:E54');
  ws.getCell('D1').value='기고 수금내역';
  ws.getCell('D1').font={name:'맑은 고딕',size:22};
  ws.getCell('D1').alignment={horizontal:'center',vertical:'middle'};
  ws.getCell('H1').value='현금'; ws.getCell('H2').value='어음'; ws.getCell('H3').value='최종 합계';
  ['H1','H2','H3'].forEach(a=>{ws.getCell(a).font={name:'맑은 고딕',size:9};ws.getCell(a).alignment={horizontal:'center',vertical:'middle'}});
  ['A','B','C','D','E'].forEach((l,i)=>{});
  const headers=['도매상','도매상명','실납코드','실납처명','채널','10월 수금대상액','최종 수금액'];
  headers.forEach((h,i)=>{
    const c=ws.getCell(6,i+1); c.value=h; c.border=borderAll(); c.font={name:'맑은 고딕',size:i>=5?10:9,bold:i===6}; c.alignment={horizontal:'center',vertical:'middle'}; if(i>=5)c.numFmt=NF;
  });
  T.forEach((x,i)=>{
    const r=POS[i];
    const vals=[10093445,'주식회사기고',Number(x.code),x.name,x.ch,BASE10[i][0],BASE10[i][1]];
    vals.forEach((v,j)=>{
      const c=ws.getCell(r,j+1); c.value=v; c.border=borderAll(); c.font={name:'맑은 고딕',size:j>=5?11:9,bold:j===6}; c.alignment={horizontal:(j===0||j===2||j===4)?'center':undefined,vertical:'middle'}; if(j>=5)c.numFmt=NF;
    });
    if(x.group==='도입현금'){ws.getCell(r,6).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFCCFFCC'}};ws.getCell(r,7).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFCCFFCC'}};}
    if(x.group==='도입어음'){ws.getCell(r,6).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFCCFFFF'}};ws.getCell(r,7).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFCCFFFF'}};}
  });
  const labels={16:'약국 현금',36:'병원 현금',45:'도입 현금',52:'도입 어음'};
  Object.entries(labels).forEach(([rr,label])=>{
    const r=Number(rr); ws.getCell(r,1).value='합계'; ws.getCell(r,2).value=label;
    for(let c=1;c<=7;c++){const x=ws.getCell(r,c);x.border=borderAll();x.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFFFFF00'}};x.font={name:'맑은 고딕',size:c>=6?11:9,bold:c===7};if(c>=6)x.numFmt=NF;}
  });
  const sums={약국현금:[0,0],병원현금:[0,0],도입현금:[0,0],도입어음:[0,0]};
  T.forEach((x,i)=>{sums[x.group][0]+=BASE10[i][0];sums[x.group][1]+=BASE10[i][1]});
  Object.entries(GROUP_ROWS).forEach(([g,r])=>{ws.getCell(r,6).value=sums[g][0];ws.getCell(r,7).value=sums[g][1]});
  ws.getCell('F1').value={formula:'F16+F36+F45'}; ws.getCell('G1').value={formula:'G16+G36+G45'};
  ws.getCell('F2').value={formula:'F52'}; ws.getCell('G2').value={formula:'G52'};
  ws.getCell('F3').value={formula:'F1+F2'}; ws.getCell('G3').value={formula:'G1+G2'};
  for(let r=1;r<=3;r++){for(let c=6;c<=7;c++){const x=ws.getCell(r,c);x.numFmt=NF;x.font={name:'맑은 고딕',size:11,bold:c===7};if(c===7)x.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFFFFF00'}};}}
  ws.pageSetup={orientation:'portrait',fitToPage:true,fitToWidth:1,fitToHeight:0,printArea:'A1:N54'};
  return wb;
}
function monthCols(month){
  const [ys,ms]=month.split('-'),y=Number(ys),m=Number(ms),diff=(y-2026)*12+(m-10);
  if(diff<0)throw new Error('2026년 10월 이후 월만 지원합니다.');
  return {target:6+diff*3,final:7+diff*3,spacer:8+diff*3,year:y,month:m};
}
function styleMonthPair(ws,target,final,spacer){
  ws.getColumn(target).width=14; ws.getColumn(final).width=17.09765625; ws.getColumn(spacer).width=14;
  for(let r=1;r<=54;r++){
    ws.getCell(r,target).style=cloneStyle(ws.getCell(r,6).style);
    ws.getCell(r,final).style=cloneStyle(ws.getCell(r,7).style);
    ws.getCell(r,spacer).style=cloneStyle(ws.getCell(r,8).style);
  }
}
function applyMonth(ws,c,month){
  const mc=monthCols(month),tc=mc.target,fc=mc.final,sc=mc.spacer;
  if(mc.month===10&&mc.year===2026){
    POS.forEach((r,i)=>{ws.getCell(r,tc).value=c.rows[i].amount;ws.getCell(r,fc).value=c.rows[i].final});
  }else{
    styleMonthPair(ws,tc,fc,sc);
    ws.getCell(6,tc).value=mc.month+'월 수금대상액'; ws.getCell(6,fc).value='최종 수금액';
    ws.getCell(1,sc).value='현금'; ws.getCell(2,sc).value='어음'; ws.getCell(3,sc).value='최종 합계';
    POS.forEach((r,i)=>{ws.getCell(r,tc).value=c.rows[i].amount;ws.getCell(r,fc).value=c.rows[i].final});
  }
  Object.entries(GROUP_ROWS).forEach(([g,r])=>{
    const arr=c.rows.filter(x=>x.group===g);
    ws.getCell(r,tc).value=arr.reduce((s,x)=>s+x.amount,0);
    ws.getCell(r,fc).value=arr.reduce((s,x)=>s+x.final,0);
  });
  const tl=ws.getColumn(tc).letter,fl=ws.getColumn(fc).letter;
  const cashRaw=['약국현금','병원현금','도입현금'].reduce((q,g)=>q+c.rows.filter(x=>x.group===g).reduce((a,x)=>a+x.amount,0),0);
  const cashFinal=['약국현금','병원현금','도입현금'].reduce((q,g)=>q+c.rows.filter(x=>x.group===g).reduce((a,x)=>a+x.final,0),0);
  const noteRaw=c.rows.filter(x=>x.group==='도입어음').reduce((a,x)=>a+x.amount,0);
  const noteFinal=c.rows.filter(x=>x.group==='도입어음').reduce((a,x)=>a+x.final,0);
  ws.getCell(1,tc).value={formula:tl+'16+'+tl+'36+'+tl+'45',result:cashRaw}; ws.getCell(1,fc).value={formula:fl+'16+'+fl+'36+'+fl+'45',result:cashFinal};
  ws.getCell(2,tc).value={formula:tl+'52',result:noteRaw}; ws.getCell(2,fc).value={formula:fl+'52',result:noteFinal};
  ws.getCell(3,tc).value={formula:tl+'1+'+tl+'2',result:cashRaw+noteRaw}; ws.getCell(3,fc).value={formula:fl+'1+'+fl+'2',result:cashFinal+noteFinal};
  return mc;
}
function openDb(){
  return new Promise((resolve,reject)=>{
    const req=indexedDB.open('gigo-collection-v3',1);
    req.onupgradeneeded=()=>{if(!req.result.objectStoreNames.contains('files'))req.result.createObjectStore('files')};
    req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error);
  });
}
async function saveCumulative(buffer,month){
  const db=await openDb();
  await new Promise((resolve,reject)=>{const tx=db.transaction('files','readwrite');tx.objectStore('files').put({month,buffer},'latest');tx.oncomplete=resolve;tx.onerror=()=>reject(tx.error)});
  db.close();
}
async function loadCumulative(){
  const db=await openDb();
  const data=await new Promise((resolve,reject)=>{const tx=db.transaction('files','readonly'),req=tx.objectStore('files').get('latest');req.onsuccess=()=>resolve(req.result||null);req.onerror=()=>reject(req.error)});
  db.close();return data;
}
async function loadBaseWorkbook(selectedMonth){
  const sharedMonths=Object.keys(gigoShared.months||{}).filter(m=>m<selectedMonth).sort();
  if(sharedMonths.length){
    const wb=buildBaseWorkbook(),ws=wb.getWorksheet('수금내역서')||wb.worksheets[0];
    for(const m of sharedMonths){
      const snap=gigoShared.months[m];
      applyMonth(ws,{rows:snap.rows},m);
    }
    return{wb,source:'공용 누적 데이터('+sharedMonths.at(-1)+')'};
  }
  const saved=await loadCumulative().catch(()=>null);
  if(saved&&saved.buffer&&saved.month&&saved.month<selectedMonth){
    try{const wb=new ExcelJS.Workbook();await wb.xlsx.load(saved.buffer);return{wb,source:'이 브라우저 누적 파일('+saved.month+')'}}catch(e){console.warn('saved workbook ignored',e)}
  }
  return{wb:buildBaseWorkbook(),source:'10월 기본 샘플'};
}

function cellText(cell){
  const v=cell.value;
  if(v==null)return '';
  if(typeof v==='object'&&v.formula!==undefined)return v.result==null?'':String(v.result);
  if(typeof v==='object'&&v.result!==undefined)return String(v.result);
  if(typeof v==='number')return Math.round(v).toLocaleString('ko-KR');
  return String(v);
}
function argbToCss(argb){
  const x=String(argb||'').replace('#','');
  if(x.length===8)return '#'+x.slice(2);
  if(x.length===6)return '#'+x;
  return null;
}

function previewMerges(ws){
  const out=[];
  try{
    const src=ws._merges||{};
    for(const m of Object.values(src)){
      if(m&&Number.isFinite(m.top)&&Number.isFinite(m.left))out.push({top:m.top,left:m.left,bottom:m.bottom,right:m.right});
      else if(m&&m.model)out.push({top:m.model.top,left:m.model.left,bottom:m.model.bottom,right:m.model.right});
    }
  }catch(_){}
  return out;
}
function previewMergeAt(merges,r,c){
  for(const m of merges)if(r>=m.top&&r<=m.bottom&&c>=m.left&&c<=m.right)return m;
  return null;
}
function previewCellText(cell){
  const v=cell.value;
  if(v==null)return '';
  if(typeof v==='object'&&v.formula!==undefined)return v.result==null?'':String(v.result);
  if(typeof v==='object'&&v.result!==undefined)return String(v.result);
  if(typeof v==='object'&&v.text!==undefined)return String(v.text);
  if(typeof v==='number')return Math.round(v).toLocaleString('ko-KR');
  return String(v);
}
function findPreviewLastRow(ws){
  let last=6;
  for(let r=6;r<=200;r++){
    let has=false;
    for(let c=1;c<=7;c++){
      const t=previewCellText(ws.getCell(r,c));
      if(String(t||'').trim()!==''){has=true;break}
    }
    if(has)last=r;
  }
  return last;
}
function previewDisplayText(cell,r,c){
  let text=previewCellText(cell);
  if(text==null)return '';
  text=String(text);
  if(c===1||c===3)return text.replace(/,/g,'');
  if((c===6||c===7)&&text){
    const raw=text.replace(/,/g,'');
    if(/^-?\d+(\.\d+)?$/.test(raw)){
      const n=Number(raw);
      if(Number.isFinite(n))return Math.round(n).toLocaleString('ko-KR');
    }
  }
  return text;
}
function previewAlign(cell,r,c){
  const a=cell.alignment?.horizontal;
  if(a==='center'||a==='right'||a==='left')return a;
  if(r===6)return 'center';
  if(c>=6)return 'right';
  if(c===1||c===3||c===5)return 'center';
  return 'left';
}
function drawPreviewText(ctx,text,x,y,w,h,align,fontSize,bold){
  if(!text)return;
  ctx.font=(bold?'700 ':'400 ')+fontSize+'px Arial, sans-serif';
  ctx.fillStyle='#111';
  ctx.textBaseline='middle';
  ctx.textAlign=align;
  const pad=6,maxW=Math.max(8,w-pad*2);
  let t=String(text);
  while(t.length>2&&ctx.measureText(t).width>maxW)t=t.slice(0,-2)+'…';
  const tx=align==='right'?x+w-pad:align==='center'?x+w/2:x+pad;
  ctx.fillText(t,tx,y+h/2);
}
async function renderExcelPreview(calcResult,month){
  try{
    const base=await loadBaseWorkbook(month),wb=base.wb,ws=wb.getWorksheet('수금내역서')||wb.worksheets[0];
    applyMonth(ws,calcResult,month);

    const firstCol=1,lastCol=7,firstRow=6,lastRow=findPreviewLastRow(ws);
    const widths=[],xPos=[0],rowHeights=[],yPos=[0];

    for(let c=firstCol;c<=lastCol;c++){
      const px=Math.max(54,Math.min(320,(ws.getColumn(c).width||12)*7.2));
      widths.push(px);xPos.push(xPos.at(-1)+px);
    }
    for(let r=firstRow;r<=lastRow;r++){
      const px=Math.max(18,Math.min(42,ws.getRow(r).height||22));
      rowHeights.push(px);yPos.push(yPos.at(-1)+px);
    }

    const totalW=xPos.at(-1),totalH=yPos.at(-1),scale=1.8;
    const canvas=$('previewCanvas'),ctx=canvas.getContext('2d');
    canvas.width=Math.round(totalW*scale);canvas.height=Math.round(totalH*scale);
    ctx.setTransform(scale,0,0,scale,0,0);
    ctx.fillStyle='#fff';ctx.fillRect(0,0,totalW,totalH);

    const merges=previewMerges(ws),skip=new Set();
    for(let r=firstRow;r<=lastRow;r++){
      for(let c=firstCol;c<=lastCol;c++){
        const k=r+':'+c;if(skip.has(k))continue;
        const m=previewMergeAt(merges,r,c);
        if(m&&(m.top!==r||m.left!==c))continue;

        let x=xPos[c-firstCol],y=yPos[r-firstRow],w=widths[c-firstCol],h=rowHeights[r-firstRow];
        if(m){
          const l=Math.max(m.left,firstCol),rr=Math.min(m.right,lastCol),t=Math.max(m.top,firstRow),bb=Math.min(m.bottom,lastRow);
          if(l<=rr&&t<=bb){
            x=xPos[l-firstCol];y=yPos[t-firstRow];
            w=xPos[rr-firstCol+1]-x;h=yPos[bb-firstRow+1]-y;
            for(let mr=t;mr<=bb;mr++)for(let mc=l;mc<=rr;mc++)if(!(mr===r&&mc===c))skip.add(mr+':'+mc);
          }
        }

        const cell=ws.getCell(r,c);
        let fill='#ffffff';
        const argb=cell.fill?.fgColor?.argb;
        if(argb)fill='#'+String(argb).slice(-6);
        else if(r===6)fill='#f1f3f5';
        ctx.fillStyle=fill;ctx.fillRect(x,y,w,h);
        ctx.strokeStyle='#222';ctx.lineWidth=.7;ctx.strokeRect(x,y,w,h);

        let text=previewDisplayText(cell,r,c);
        if((c===6||c===7)&&String(text).trim()==='0')text='-';
        const align=previewAlign(cell,r,c);
        const fontSize=Math.max(8,Math.min(14,cell.font?.size||9));
        drawPreviewText(ctx,text,x,y,w,h,align,fontSize,!!cell.font?.bold);
      }
    }

    const png=canvas.toDataURL('image/png');
    previewJpgUrl=canvas.toDataURL('image/jpeg',0.96);
    $('previewImage').src=png;
    $('previewShell').style.display='flex';$('previewEmpty').style.display='none';
    $('previewSub').textContent=month+' 기준 · 6행부터 G열까지 전체 미리보기';
    if($('saveJpgBtn'))$('saveJpgBtn').disabled=false;
  }catch(e){
    console.error('preview error',e);
    $('previewSub').textContent='미리보기 생성 중 오류가 발생했습니다.';
    if($('saveJpgBtn'))$('saveJpgBtn').disabled=true;
  }
}
$('saveJpgBtn')?.addEventListener('click',()=>{
  if(!previewJpgUrl)return;
  const month=$('targetMonth').value||'';
  const [yy,mm]=month.split('-');
  const a=document.createElement('a');
  a.href=previewJpgUrl;
  a.download='기고 수금내역_'+(yy||'')+'-'+(mm||'')+'_미리보기.jpg';
  a.click();
});
$('masterLoginBtn')?.addEventListener('click',masterLogin);
$('masterRefreshBtn')?.addEventListener('click',()=>pullSharedData(true));
$('masterSaveBtn')?.addEventListener('click',saveMasterCurrent);
$('masterPassword')?.addEventListener('keydown',e=>{if(e.key==='Enter')masterLogin()});
(async()=>{
  setMasterUI(false);
  const savedPw=sessionStorage.getItem('gigo-master-password');
  if(savedPw&&syncEndpoint()){
    try{await syncRequest({type:'gigoMasterLogin',password:savedPw});setMasterUI(true);masterStatus('마스터 세션 복원 완료','good')}catch(_){sessionStorage.removeItem('gigo-master-password')}
  }
  await pullSharedData(true);
})();
$('downloadBtn').addEventListener('click',async()=>{
  try{
    if(!analysis)return;
    $('downloadBtn').disabled=true;
    const month=$('targetMonth').value||'';
    status('샘플 양식에 '+month+' 데이터를 누적하고 있습니다.');
    const base=await loadBaseWorkbook(month),wb=base.wb,ws=wb.getWorksheet('수금내역서')||wb.worksheets[0];
    if(!ws)throw new Error('수금내역서 시트를 생성할 수 없습니다.');
    const mc=applyMonth(ws,analysis.calc,month);
    const yy=String(mc.year),mm=String(mc.month).padStart(2,'0');
    const dcmName=yy.slice(2)+'.'+mm+' DCM수금내역서',sapName=yy.slice(2)+'.'+mm+' SAP매출';
    const oldD=wb.getWorksheet(dcmName); if(oldD)wb.removeWorksheet(oldD.id);
    const oldS=wb.getWorksheet(sapName); if(oldS)wb.removeWorksheet(oldS.id);
    copy(analysis.dws,wb,dcmName);
    const ss=copy(analysis.sws,wb,sapName);
    ss.getCell(1,75).value='간납가';ss.getCell(1,77).value='유형';
    for(let r=2;r<=analysis.sws.rowCount;r++){
      ss.getCell(r,75).value=num(val(analysis.sws,r,49))+num(val(analysis.sws,r,52));
      const p=String(val(analysis.sws,r,25)||'');ss.getCell(r,77).value=(p.includes('콩코르')||p.includes('크레스토'))?'현금':'어음';
    }
    const buf=await wb.xlsx.writeBuffer();
    await saveCumulative(buf,month).catch(()=>{});
    const blob=new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),a=document.createElement('a');
    a.href=URL.createObjectURL(blob);a.download='기고 수금 내역서_'+yy+'년 '+mm+'월_누적.xlsx';a.click();URL.revokeObjectURL(a.href);
    status('완료: '+base.source+' 기준으로 '+mc.month+'월 데이터가 '+ws.getColumn(mc.target).letter+'/'+ws.getColumn(mc.final).letter+'열에 저장되었습니다.','good');
  }catch(e){console.error(e);status('생성 오류: '+e.message,'bad')}
  finally{$('downloadBtn').disabled=false}
});