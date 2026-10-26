
const $=id=>document.getElementById(id);
let dcmFile=null,sapFile=null,analysis=null;
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
function wire(i,d,n,set){const I=$(i),D=$(d),N=$(n),use=f=>{if(!f)return;if(!/\.xlsx$/i.test(f.name)){N.textContent='xlsx 파일만 지원합니다.';N.className='fileline bad';return}set(f);N.textContent='✓ '+f.name;N.className='fileline';ready()};I.addEventListener('change',e=>use(e.target.files[0]));['dragenter','dragover'].forEach(x=>D.addEventListener(x,e=>{e.preventDefault();D.classList.add('drag')}));['dragleave','drop'].forEach(x=>D.addEventListener(x,e=>{e.preventDefault();D.classList.remove('drag')}));D.addEventListener('drop',e=>use(e.dataTransfer.files[0]))}
wire('dcmFile','dcmDrop','dcmName',f=>dcmFile=f);wire('sapFile','sapDrop','sapName',f=>sapFile=f);
function ready(){$('analyzeBtn').disabled=!(dcmFile&&sapFile);$('downloadBtn').disabled=true;analysis=null}
const now=new Date();$('targetMonth').value=now.getFullYear()+'-'+String(now.getMonth()+1).padStart(2,'0');
async function load(f){const w=new ExcelJS.Workbook();await w.xlsx.load(await f.arrayBuffer());return w}
function sheet(w,k){if(k==='dcm')return w.worksheets.find(s=>String(val(s,5,2)).includes('실납코드')&&String(val(s,5,4)).includes('채널'))||w.worksheets[0];return w.worksheets.find(s=>String(val(s,1,17)).includes('실납처')&&String(val(s,1,22)).includes('채널'))||w.worksheets[0]}
function dmap(ws){const p=new Map(),t=new Map();let rows=0;for(let r=8;r<=ws.rowCount;r++){const c=code(val(ws,r,2)),h=String(val(ws,r,4)||'').trim();if(!c||!h)continue;rows++;const k=c+'|'+h;p.set(k,(p.get(k)||0)+num(val(ws,r,12)));t.set(k,(t.get(k)||0)+num(val(ws,r,14)))}return{prev:p,total:t,rows}}
function smap(ws){const cash=new Map(),note=new Map();let rows=0;for(let r=2;r<=ws.rowCount;r++){const c=code(val(ws,r,17)),h=String(val(ws,r,22)||'').trim();if(!c||!h)continue;const pr=String(val(ws,r,25)||''),cashFlag=pr.includes('콩코르')||pr.includes('크레스토'),amt=num(val(ws,r,49))+num(val(ws,r,52)),k=c+'|'+h,m=cashFlag?cash:note;rows++;m.set(k,(m.get(k)||0)+amt)}return{cash,note,rows}}
function calc(d,s){const rows=[];let matched=0,unmatched=0;for(const x of T){const k=x.code+'|'+x.ch,exc=(x.ch==='DW24'&&EX.has(x.code))||(x.code==='10093445'&&(x.ch==='DW21'||x.ch==='DW22'));let amt=0,src='';if(exc){amt=d.total.get(k)||0;src='DCM 합계'}else if(SCH.has(x.ch)){amt=(x.payType==='현금'?s.cash:s.note).get(k)||0;src='SAP '+x.payType}else{amt=d.prev.get(k)||0;src='DCM 전월'}const found=amt!==0||(exc?d.total.has(k):SCH.has(x.ch)?(x.payType==='현금'?s.cash:s.note).has(k):d.prev.has(k));found?matched++:unmatched++;rows.push({...x,amount:amt,final:amt,source:src,found,isExc:exc})}for(const g of G){const a=rows.filter(x=>x.group===g),sum=a.reduce((q,x)=>q+x.amount,0),diff=(sum<=0?0:Math.ceil(sum/1000000)*1000000)-sum;if(diff>0&&a.length){let m=a[0];for(const x of a)if(x.amount>m.amount)m=x;m.final+=diff}}return{rows,matched,unmatched}}
function show(a,dr,sr){$('summary').style.display='grid';$('resultCard').style.display='block';$('mDcm').textContent=dr.toLocaleString();$('mSap').textContent=sr.toLocaleString();const z={};G.forEach(g=>z[g]=a.rows.filter(x=>x.group===g).reduce((q,x)=>q+x.final,0));$('m1').textContent=fmt(z['약국현금']);$('m2').textContent=fmt(z['병원현금']);$('m3').textContent=fmt(z['도입현금']);$('m4').textContent=fmt(z['도입어음']);$('matchInfo').textContent='매칭 '+a.matched+'건 · 미매칭 '+a.unmatched+'건';$('resultBody').innerHTML=a.rows.map(x=>'<tr><td>'+x.group+'</td><td>'+x.name+'</td><td>'+x.ch+'</td><td class="num">'+fmt(x.amount)+'</td><td class="num">'+fmt(x.final)+'</td><td><span class="tag '+(!x.found?'none':x.isExc?'exc':x.source.startsWith('SAP')?'sap':'dcm')+'">'+(!x.found?'미매칭':x.source)+'</span></td></tr>').join('')}
$('analyzeBtn').addEventListener('click',async()=>{try{$('analyzeBtn').disabled=true;status('파일 구조를 확인하고 수금대상을 계산하고 있습니다.');const [dw,sw]=await Promise.all([load(dcmFile),load(sapFile)]),ds=sheet(dw,'dcm'),ss=sheet(sw,'sap');if(!ds||!ss)throw new Error('필수 시트를 찾을 수 없습니다.');const d=dmap(ds),s=smap(ss),c=calc(d,s);analysis={calc:c,dws:ds,sws:ss};show(c,d.rows,s.rows);$('downloadBtn').disabled=false;status(c.unmatched?'계산 완료: '+c.matched+'건 매칭, '+c.unmatched+'건 미매칭입니다.':'계산 완료: 대상 '+c.matched+'건이 모두 매칭되었습니다.',c.unmatched?'':'good')}catch(e){status('오류: '+e.message,'bad')}finally{$('analyzeBtn').disabled=false}});

function copy(src,wb,name){
  const dst=wb.addWorksheet(name);
  for(let r=1;r<=src.rowCount;r++){
    const a=[];
    for(let c=1;c<=src.columnCount;c++) a.push(val(src,r,c));
    dst.addRow(a);
  }
  return dst;
}
function cloneStyle(s){ return s ? JSON.parse(JSON.stringify(s)) : {}; }
function monthCols(month){
  const [ys,ms]=month.split('-');
  const y=Number(ys), m=Number(ms);
  const diff=(y-2026)*12+(m-10);
  if(diff<0) throw new Error('2026년 10월 이후 월만 누적 생성할 수 있습니다.');
  const target=6+diff*3;
  return {target,final:target+1,spacer:target-1,year:y,month:m};
}
function styleMonthPair(ws,target,final,spacer){
  ws.getColumn(target).width=ws.getColumn(6).width;
  ws.getColumn(final).width=ws.getColumn(7).width;
  ws.getColumn(spacer).width=ws.getColumn(8).width;
  for(let r=1;r<=54;r++){
    ws.getCell(r,target).style=cloneStyle(ws.getCell(r,6).style);
    ws.getCell(r,final).style=cloneStyle(ws.getCell(r,7).style);
    ws.getCell(r,spacer).style=cloneStyle(ws.getCell(r,8).style);
  }
}
function applyMonth(ws,c,month){
  const mc=monthCols(month), tc=mc.target, fc=mc.final, sc=mc.spacer;
  styleMonthPair(ws,tc,fc,sc);
  ws.getCell(6,tc).value=mc.month+'월 수금대상액';
  ws.getCell(6,fc).value='최종 수금액';
  ws.getCell(6,sc).value=null;
  ws.getCell(1,sc).value='현금';
  ws.getCell(2,sc).value='어음';
  ws.getCell(3,sc).value='최종 합계';

  const pos=[7,8,9,10,11,12,13,17,18,19,20,21,22,23,24,25,26,27,28,29,30,37,38,39,40,41,42,43,44,46,47,48,49,50,51];
  c.rows.forEach((x,i)=>{
    const r=pos[i];
    ws.getCell(r,tc).value=x.amount;
    ws.getCell(r,fc).value=x.final;
  });
  const groups={
    '약국현금':{row:16,start:7,end:13},
    '병원현금':{row:36,start:17,end:30},
    '도입현금':{row:45,start:37,end:44},
    '도입어음':{row:52,start:46,end:51}
  };
  Object.entries(groups).forEach(([g,o])=>{
    const arr=c.rows.filter(x=>x.group===g);
    ws.getCell(o.row,tc).value=arr.reduce((s,x)=>s+x.amount,0);
    ws.getCell(o.row,fc).value=arr.reduce((s,x)=>s+x.final,0);
  });
  const tL=ws.getColumn(tc).letter, fL=ws.getColumn(fc).letter;
  ws.getCell(1,tc).value={formula:tL+'16+'+tL+'36+'+tL+'45'};
  ws.getCell(1,fc).value={formula:fL+'16+'+fL+'36+'+fL+'45'};
  ws.getCell(2,tc).value={formula:tL+'52'};
  ws.getCell(2,fc).value={formula:fL+'52'};
  ws.getCell(3,tc).value={formula:tL+'1+'+tL+'2'};
  ws.getCell(3,fc).value={formula:fL+'1+'+fL+'2'};
  return mc;
}
function openDb(){
  return new Promise((resolve,reject)=>{
    const req=indexedDB.open('gigo-collection',1);
    req.onupgradeneeded=()=>{ if(!req.result.objectStoreNames.contains('files')) req.result.createObjectStore('files'); };
    req.onsuccess=()=>resolve(req.result); req.onerror=()=>reject(req.error);
  });
}
async function saveCumulative(buffer,month){
  const db=await openDb();
  await new Promise((resolve,reject)=>{
    const tx=db.transaction('files','readwrite');
    tx.objectStore('files').put({month,buffer},'latest');
    tx.oncomplete=resolve; tx.onerror=()=>reject(tx.error);
  });
  db.close();
}
async function loadCumulative(){
  const db=await openDb();
  const data=await new Promise((resolve,reject)=>{
    const tx=db.transaction('files','readonly'),req=tx.objectStore('files').get('latest');
    req.onsuccess=()=>resolve(req.result||null); req.onerror=()=>reject(req.error);
  });
  db.close(); return data;
}
async function loadBaseWorkbook(selectedMonth){
  const saved=await loadCumulative().catch(()=>null);
  let buf=null, source='10월 기본 템플릿';
  if(saved && saved.buffer){
    const s=String(saved.month||'');
    if(s && s<=selectedMonth){ buf=saved.buffer; source='이전 누적 파일('+s+')'; }
  }
  if(!buf){
    const res=await fetch('template.xlsx?v=20261026-5',{cache:'no-store'});
    if(!res.ok) throw new Error('기본 샘플 템플릿을 불러오지 못했습니다.');
    buf=await res.arrayBuffer();
  }
  const wb=new ExcelJS.Workbook(); await wb.xlsx.load(buf); return {wb,source};
}
$('downloadBtn').addEventListener('click',async()=>{
  try{
    if(!analysis)return;
    $('downloadBtn').disabled=true;
    const month=$('targetMonth').value||'';
    status('샘플 양식을 불러와 '+month+' 데이터를 누적하고 있습니다.');
    const base=await loadBaseWorkbook(month), wb=base.wb;
    let ws=wb.getWorksheet('수금내역서')||wb.worksheets[0];
    if(!ws) throw new Error('샘플 템플릿의 수금내역서 시트를 찾을 수 없습니다.');
    const mc=applyMonth(ws,analysis.calc,month);

    wb.worksheets.slice().forEach(s=>{ if(s.id!==ws.id) wb.removeWorksheet(s.id); });
    const yy=String(mc.year), mm=String(mc.month).padStart(2,'0');
    copy(analysis.dws,wb,yy.slice(2)+'.'+mm+' DCM수금내역서');
    const ss=copy(analysis.sws,wb,yy.slice(2)+'.'+mm+' SAP매출');
    ss.getCell(1,75).value='간납가'; ss.getCell(1,77).value='유형';
    for(let r=2;r<=analysis.sws.rowCount;r++){
      ss.getCell(r,75).value=num(val(analysis.sws,r,49))+num(val(analysis.sws,r,52));
      const p=String(val(analysis.sws,r,25)||'');
      ss.getCell(r,77).value=(p.includes('콩코르')||p.includes('크레스토'))?'현금':'어음';
    }
    const buf=await wb.xlsx.writeBuffer();
    await saveCumulative(buf,month).catch(()=>{});
    const blob=new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),a=document.createElement('a');
    a.href=URL.createObjectURL(blob); a.download='기고 수금 내역서_'+yy+'년 '+mm+'월_누적.xlsx'; a.click(); URL.revokeObjectURL(a.href);
    status('완료: '+base.source+'을 기준으로 '+mc.month+'월 값이 '+ws.getColumn(mc.target).letter+'/'+ws.getColumn(mc.final).letter+'열에 누적되었습니다.','good');
  }catch(e){ console.error(e); status('생성 오류: '+e.message,'bad'); }
  finally{$('downloadBtn').disabled=false;}
});