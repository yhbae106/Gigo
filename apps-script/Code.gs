const GIGO_FILE_PROP='GIGO_SHARED_FILE_ID';
const GIGO_MASTER_PROP='GIGO_MASTER_PASSWORD';
const GIGO_FILE_NAME='Gigo_Master_Data.json';

function doPost(e){
  try{
    const body=JSON.parse((e&&e.postData&&e.postData.contents)||'{}');
    if(body.type==='gigoLoad'){
      const x=loadGigo_();
      return json_({ok:true,data:x.data,updatedAt:x.updatedAt});
    }
    if(body.type==='gigoMasterLogin'){
      assertGigoMaster_(body.password);
      return json_({ok:true});
    }
    if(body.type==='gigoSave'){
      assertGigoMaster_(body.password);
      const month=String(body.month||'').trim();
      if(!/^20\d{2}-\d{2}$/.test(month))throw new Error('저장 월 형식이 올바르지 않습니다.');
      const snap=body.snapshot;
      if(!snap||!Array.isArray(snap.rows))throw new Error('저장 데이터가 없습니다.');
      const current=loadGigo_().data||{version:1,months:{}};
      if(!current.months||typeof current.months!=='object')current.months={};
      current.months[month]={
        rows:snap.rows,
        matched:Number(snap.matched||0),
        unmatched:Number(snap.unmatched||0),
        savedAt:new Date().toISOString()
      };
      current.version=1;
      const updatedAt=new Date().toISOString();
      saveGigo_(current,updatedAt);
      return json_({ok:true,data:current,updatedAt:updatedAt});
    }
    throw new Error('지원하지 않는 요청입니다.');
  }catch(err){
    return json_({ok:false,error:String(err&&err.message||err)});
  }
}
function assertGigoMaster_(pw){
  const expected=PropertiesService.getScriptProperties().getProperty(GIGO_MASTER_PROP);
  if(!expected)throw new Error('GIGO_MASTER_PASSWORD가 설정되지 않았습니다.');
  if(!pw||String(pw)!==expected)throw new Error('마스터 비밀번호가 올바르지 않습니다.');
}
function loadGigo_(){
  const props=PropertiesService.getScriptProperties();
  const id=props.getProperty(GIGO_FILE_PROP);
  if(!id)return {data:{version:1,months:{}},updatedAt:''};
  try{
    const f=DriveApp.getFileById(id);
    const text=f.getBlob().getDataAsString('UTF-8');
    const parsed=JSON.parse(text||'{}');
    return {data:parsed.data||{version:1,months:{}},updatedAt:parsed.updatedAt||''};
  }catch(e){
    props.deleteProperty(GIGO_FILE_PROP);
    return {data:{version:1,months:{}},updatedAt:''};
  }
}
function saveGigo_(data,updatedAt){
  const props=PropertiesService.getScriptProperties();
  const payload=JSON.stringify({updatedAt:updatedAt,data:data});
  const id=props.getProperty(GIGO_FILE_PROP);
  if(id){
    try{
      DriveApp.getFileById(id).setContent(payload);
      return;
    }catch(e){props.deleteProperty(GIGO_FILE_PROP);}
  }
  const file=DriveApp.createFile(GIGO_FILE_NAME,payload,MimeType.PLAIN_TEXT);
  props.setProperty(GIGO_FILE_PROP,file.getId());
}
function setGigoMasterPassword(){
  const ui=SpreadsheetApp.getUi();
  const res=ui.prompt('Gigo 마스터 비밀번호 설정','6자 이상 마스터 비밀번호를 입력하세요.',ui.ButtonSet.OK_CANCEL);
  if(res.getSelectedButton()!==ui.Button.OK)return;
  const pw=res.getResponseText().trim();
  if(pw.length<6)throw new Error('6자 이상으로 설정하세요.');
  PropertiesService.getScriptProperties().setProperty(GIGO_MASTER_PROP,pw);
  ui.alert('Gigo 마스터 비밀번호가 저장되었습니다.');
}
function json_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
