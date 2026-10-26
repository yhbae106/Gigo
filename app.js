
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
function makeSharedSheet(wb,c,month){
  const ws=wb.addWorksheet('수금내역서',{views:[{showGridLines:false}]});
  // 원본 공유용 구조: A:V 숨김, W:AD 출력영역
  for(let i=1;i<=22;i++) ws.getColumn(i).hidden=true;
  const widths={23:12.199,24:10.598,25:10.199,26:22.098,27:9.598,28:14.0,29:17.098,30:14.0};
  Object.entries(widths).forEach(([i,w])=>ws.getColumn(Number(i)).width=w);
  for(let r=1;r<=56;r++) ws.getRow(r).height=22.8;
  [14,15,31,32,33,34,35].forEach(r=>ws.getRow(r).hidden=true);

  const thin={style:'thin',color:{argb:'FF000000'}};
  const border={top:thin,left:thin,bottom:thin,right:thin};
  const numFmt='_-* #,##0_-;\\-* #,##0_-;_-* "-"_-;_-@_-';
  const normal={name:'맑은 고딕',size:9};
  const amountFont={name:'맑은 고딕',size:11};
  const yellow={type:'pattern',pattern:'solid',fgColor:{argb:'FFFFFF00'}};
  const green={type:'pattern',pattern:'solid',fgColor:{argb:'FFCCFFCC'}};
  const cyan={type:'pattern',pattern:'solid',fgColor:{argb:'FFCCFFFF'}};

  ws.mergeCells('Z1:AA3');
  ws.getCell('Z1').value='기고 수금내역';
  ws.getCell('Z1').font={name:'맑은 고딕',size:22};
  ws.getCell('Z1').alignment={horizontal:'center',vertical:'middle'};

  const cashRows=c.rows.filter(x=>x.group!=='도입어음');
  const noteRows=c.rows.filter(x=>x.group==='도입어음');
  const cashRaw=cashRows.reduce((s,x)=>s+x.amount,0), cashFinal=cashRows.reduce((s,x)=>s+x.final,0);
  const noteRaw=noteRows.reduce((s,x)=>s+x.amount,0), noteFinal=noteRows.reduce((s,x)=>s+x.final,0);
  [[cashRaw,cashFinal],[noteRaw,noteFinal],[cashRaw+noteRaw,cashFinal+noteFinal]].forEach((v,i)=>{
    const r=i+1;
    ws.getCell(r,28).value=v[0]; // AB
    ws.getCell(r,29).value=v[1]; // AC
    ws.getCell(r,28).numberFormat=numFmt;
    ws.getCell(r,29).numberFormat=numFmt;
    ws.getCell(r,28).font=amountFont;
    ws.getCell(r,29).font={...amountFont,bold:true};
    ws.getCell(r,29).fill=yellow;
  });

  const mm=(month.split('-')[1]||'').replace(/^0/,'');
  const headers=['도매상','도매상명','실납코드','실납처명','채널',(mm||'')+'월 수금대상액','최종 수금액'];
  for(let i=0;i<headers.length;i++){
    const cell=ws.getCell(6,23+i);
    cell.value=headers[i];
    cell.font={...normal,bold:i===6};
    cell.alignment={horizontal:'center',vertical:'middle'};
    cell.border=border;
    if(i>=5) cell.numberFormat=numFmt;
  }

  const pos=[7,8,9,10,11,12,13,17,18,19,20,21,22,23,24,25,26,27,28,29,30,37,38,39,40,41,42,43,44,46,47,48,49,50,51];
  c.rows.forEach((x,i)=>{
    const r=pos[i], vals=[10093445,'주식회사기고',Number(x.code),x.name,x.ch,x.amount,x.final];
    vals.forEach((v,j)=>{
      const cell=ws.getCell(r,23+j);
      cell.value=v;
      cell.border=border;
      cell.font=j>=5?(j===6?{...amountFont,bold:true}:amountFont):normal;
      cell.alignment={vertical:'middle',horizontal:(j===0||j===2||j===4)?'center':undefined};
      if(j>=5) cell.numberFormat=numFmt;
    });
    if(x.group==='도입현금'){ws.getCell(r,28).fill=green;ws.getCell(r,29).fill=green;}
    if(x.group==='도입어음'){ws.getCell(r,28).fill=cyan;ws.getCell(r,29).fill=cyan;}
  });

  // 원본의 빈 행도 테두리/숫자서식 유지
  for(let r=6;r<=52;r++){
    for(let cidx=23;cidx<=29;cidx++){
      const cell=ws.getCell(r,cidx);
      cell.border=border;
      if(cidx>=28) cell.numberFormat=numFmt;
    }
  }

  const totals={
    '약국현금':{row:16,type:'현금'},
    '병원현금':{row:36,type:'현금'},
    '도입현금':{row:45,type:'현금'},
    '도입어음':{row:52,type:'어음'}
  };
  Object.entries(totals).forEach(([g,info])=>{
    const arr=c.rows.filter(x=>x.group===g), raw=arr.reduce((s,x)=>s+x.amount,0), fin=arr.reduce((s,x)=>s+x.final,0), r=info.row;
    ws.getCell(r,23).value='합계';
    ws.getCell(r,24).value=info.type;
    ws.getCell(r,28).value=raw;
    ws.getCell(r,29).value=fin;
    ws.getCell(r,30).value=g;
    for(let ci=23;ci<=29;ci++){
      const cell=ws.getCell(r,ci);
      cell.fill=yellow;
      cell.border=border;
      cell.font=ci===29?{...amountFont,bold:true}:normal;
      if(ci>=28) cell.numberFormat=numFmt;
    }
    ws.getCell(r,30).font=normal;
  });

  // 원본처럼 출력영역만 보이도록
  ws.pageSetup.orientation='portrait';
  ws.pageSetup.fitToPage=true;
  ws.pageSetup.fitToWidth=1;
  ws.pageSetup.fitToHeight=0;
  ws.pageSetup.printArea='W1:AD52';
  ws.sheetView.showGridLines=false;
  return ws;
}
$('downloadBtn').addEventListener('click',async()=>{
  try{
    if(!analysis)return;
    $('downloadBtn').disabled=true;
    status('공유용 양식으로 결과 엑셀을 생성하고 있습니다.');
    const month=$('targetMonth').value||'',a=month.split('-'),yy=a[0]||'',mm=a[1]||'';
    const wb=new ExcelJS.Workbook();
    wb.creator='Gigo Collection Dashboard';
    makeSharedSheet(wb,analysis.calc,month);

    const dcmName=yy.slice(2)+'.'+mm+' DCM수금내역서';
    const sapName=yy.slice(2)+'.'+mm+' SAP매출';
    copy(analysis.dws,wb,dcmName);
    const ss=copy(analysis.sws,wb,sapName);
    ss.getCell(1,75).value='간납가';
    ss.getCell(1,77).value='유형';
    for(let r=2;r<=analysis.sws.rowCount;r++){
      ss.getCell(r,75).value=num(val(analysis.sws,r,49))+num(val(analysis.sws,r,52));
      const p=String(val(analysis.sws,r,25)||'');
      ss.getCell(r,77).value=(p.includes('콩코르')||p.includes('크레스토'))?'현금':'어음';
    }

    const buf=await wb.xlsx.writeBuffer();
    const blob=new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
    const ael=document.createElement('a');
    ael.href=URL.createObjectURL(blob);
    ael.download='기고 수금 내역서_'+yy+'년 '+mm+'월_공유용.xlsx';
    ael.click();
    URL.revokeObjectURL(ael.href);
    status('첨부하신 공유용 양식으로 수금내역서 생성이 완료되었습니다.','good');
  }catch(e){
    status('생성 오류: '+e.message,'bad');
  }finally{
    $('downloadBtn').disabled=false;
  }
});
