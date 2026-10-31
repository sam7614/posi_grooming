/** 교수자 소유의 새 Google Sheets에 연결해서 사용합니다. */
var NAMES_ = {people:'참여자', current:'현재기록', history:'수정이력', config:'수업설정'};
var HEADERS_ = {
 people:['참여코드','표시이름','개인키해시','상태'],
 current:['참여코드','표시이름','사례번호','버전','저장시각','관찰한 사실','해석과 추가 확인','다음 행동과 이유','수정한 점과 이유'],
 history:['요청ID','저장시각','참여코드','사례번호','버전','관찰한 사실','해석과 추가 확인','다음 행동과 이유','수정한 점과 이유'],
 config:['항목','값']
};
function onOpen(){
 SpreadsheetApp.getUi().createMenu('학습기록 관리')
  .addItem('1. 처음 설정','setup_').addItem('2. 참여자 등록·키 재발급','issueKey_')
  .addItem('현재기록 다시 만들기','rebuildCurrent_').addToUi();
}
function setup_(){
 var ss=SpreadsheetApp.getActiveSpreadsheet();
 if(!ss)throw new Error('Google Sheets의 확장 프로그램에서 이 프로젝트를 여세요.');
 var old=PropertiesService.getScriptProperties().getProperty('BOOK_ID');
 if(old && old!==ss.getId())throw new Error('다른 시트 연결이 있습니다. 새 시트에서 새 프로젝트를 만드세요.');
 var lock=LockService.getScriptLock();lock.waitLock(15000);
 try{
  Object.keys(NAMES_).forEach(function(k){
   var sh=ss.getSheetByName(NAMES_[k])||ss.insertSheet(NAMES_[k]);
   if(sh.getLastRow()===0){sh.appendRow(HEADERS_[k]);sh.setFrozenRows(1);sh.getRange(1,1,1,HEADERS_[k].length).setFontWeight('bold');}
   else if(JSON.stringify(sh.getRange(1,1,1,HEADERS_[k].length).getValues()[0])!==JSON.stringify(HEADERS_[k]))throw new Error(NAMES_[k]+'의 열 제목이 다릅니다. 기존 자료를 지우지 말고 새 시트를 사용하세요.');
  });
  PropertiesService.getScriptProperties().setProperty('BOOK_ID',ss.getId());
  var c=ss.getSheetByName(NAMES_.config);
  if(c.getLastRow()===1)c.getRange(2,1,4,2).setValues([
   ['수업명','긍정강화그루밍 관찰과 판단'],['사례번호','case-01'],
   ['사례설명','도구가 앞발 쪽으로 가까워지자 반려견이 앞발을 뒤로 옮기는 장면이 보였습니다. 전후 상황은 모두 제시되지 않았습니다. 관찰과 해석을 구분해 기록하세요.'],['접수상태','OPEN']
  ]);
  SpreadsheetApp.flush();
 }finally{lock.releaseLock();}
 SpreadsheetApp.getUi().alert('설정 완료. 시트로 돌아와 새로고침한 뒤 학습기록 관리 메뉴에서 참여자를 등록하세요.');
}
function issueKey_(){
 var ui=SpreadsheetApp.getUi();
 var c=ui.prompt('참여자 등록 또는 키 재발급','참여코드 예: A01 (영문, 숫자, 하이픈, 밑줄)',ui.ButtonSet.OK_CANCEL);
 if(c.getSelectedButton()!==ui.Button.OK)return;
 var code=c.getResponseText().trim();if(!/^[A-Za-z0-9_-]{1,40}$/.test(code))throw new Error('참여코드는 영문·숫자·하이픈·밑줄 1~40자입니다.');
 var n=ui.prompt('표시이름','교수자가 알아볼 이름이나 별칭. 기존 코드이면 키가 교체됩니다.',ui.ButtonSet.OK_CANCEL);
 if(n.getSelectedButton()!==ui.Button.OK)return;
 var name=n.getResponseText().trim().slice(0,80);if(!name)throw new Error('표시이름을 입력하세요.');
 var key=Utilities.getUuid().replace(/-/g,'')+Utilities.getUuid().replace(/-/g,'');
 var lock=LockService.getScriptLock();lock.waitLock(15000);
 try{
  var sh=book_().getSheetByName(NAMES_.people),rows=rows_(sh),found=rows.findIndex(function(r){return String(r[0])===code;});
  var target=found<0?sh.getLastRow()+1:found+2;
  sh.getRange(target,1,1,4).setValues([[code,safe_(name),hash_(key),'ACTIVE']]);SpreadsheetApp.flush();
 }finally{lock.releaseLock();}
 var html='<p>참여코드: <strong>'+code+'</strong></p><p>아래 개인키를 복사해 해당 학생에게만 전달하세요. 닫으면 원문을 다시 조회할 수 없습니다. 재발급하면 이전 키는 즉시 사용할 수 없습니다.</p><textarea readonly style="width:95%;height:85px">'+key+'</textarea><p>시트 주소나 편집 권한은 학생에게 공유하지 않습니다.</p>';
 ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(490).setHeight(300),'개인 접속키');
}
function doGet(){return HtmlService.createHtmlOutputFromFile('Index').setTitle('수업 학습기록').addMetaTag('viewport','width=device-width, initial-scale=1');}
// 웹앱에서 호출할 수 있는 학생용 함수는 아래 두 개뿐입니다.
function getRecord(auth){
 var person=authenticate_(auth),cfg=config_(),list=history_();
 var rec=latest_(list,person.code,cfg.caseId);
 return result_(person,cfg,rec);
}
function saveRecord(auth,payload){
 var lock=LockService.getScriptLock();lock.waitLock(15000);
 try{
  var person=authenticate_(auth),cfg=config_();
  if(cfg.status!=='OPEN')throw new Error('접수가 마감되었습니다. 교수자에게 문의하세요.');
  if(!payload||payload.caseId!==cfg.caseId)throw new Error('사례가 변경되었습니다. 기록을 복사한 뒤 다시 불러오세요.');
  if(!/^[A-Za-z0-9_-]{8,100}$/.test(String(payload.requestId||'')))throw new Error('요청ID가 올바르지 않습니다.');
  var fields=['fact','interpretation','action','change'];var clean={};
  fields.forEach(function(k){if(typeof payload[k]!=='string'||payload[k].length>3000)throw new Error('각 항목을 3000자 이하로 작성하세요.');clean[k]=payload[k].trim();});
  if(!clean.fact||!clean.interpretation||!clean.action)throw new Error('관찰, 해석, 다음 행동의 세 항목을 작성하세요.');
  var all=history_(),duplicate=all.find(function(r){return String(r[0])===payload.requestId;});
  if(duplicate && (String(duplicate[2])!==person.code||String(duplicate[3])!==cfg.caseId))throw new Error('요청ID가 중복되었습니다. 기록을 복사한 뒤 다시 불러오세요.');
  var rec=latest_(all,person.code,cfg.caseId);
  if(duplicate && rec && String(rec[0])!==payload.requestId)throw new Error('이 요청이 저장된 뒤 다른 기록이 저장되었습니다. 내 내용을 복사한 뒤 다시 불러와 비교하세요.');
  if(!duplicate){
   if(!Number.isInteger(payload.baseRevision)||payload.baseRevision!==(rec?Number(rec[4]):0))throw new Error('다른 창에서 기록이 변경되었습니다. 내 내용을 복사한 뒤 다시 불러와 비교하세요.');
   var next=(rec?Number(rec[4]):0)+1;
   rec=[payload.requestId,new Date().toISOString(),person.code,cfg.caseId,next,safe_(clean.fact),safe_(clean.interpretation),safe_(clean.action),safe_(clean.change)];
   book_().getSheetByName(NAMES_.history).appendRow(rec);SpreadsheetApp.flush();
  }
  // 이력 저장 후 현재기록 갱신 실패 시 같은 요청ID로 재시도하면 중복 이력 없이 복구합니다.
  updateCurrent_(person,rec);SpreadsheetApp.flush();
  return result_(person,cfg,rec);
 }finally{lock.releaseLock();}
}
function book_(){var id=PropertiesService.getScriptProperties().getProperty('BOOK_ID');if(!id)throw new Error('교수자의 처음 설정이 필요합니다.');return SpreadsheetApp.openById(id);}
function rows_(sh){return sh.getLastRow()<2?[]:sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();}
function config_(){
 var d={};rows_(book_().getSheetByName(NAMES_.config)).forEach(function(r){d[String(r[0])]=String(r[1]);});
 if(!d['사례번호'])throw new Error('수업설정의 사례번호가 없습니다.');
 return {title:d['수업명'],caseId:d['사례번호'],description:d['사례설명'],status:d['접수상태']};
}
function hash_(s){return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,s,Utilities.Charset.UTF_8).map(function(b){return ('0'+(b&255).toString(16)).slice(-2);}).join('');}
function authenticate_(auth){
 if(!auth||!/^[A-Za-z0-9_-]{1,40}$/.test(String(auth.code||''))||!/^[A-Za-z0-9_-]{6,64}$/.test(String(auth.key||'')))throw new Error('참여코드 또는 개인키를 확인하세요.');
 var h=hash_(auth.key),r=rows_(book_().getSheetByName(NAMES_.people)).find(function(r){return String(r[0])===auth.code&&String(r[2])===h&&String(r[3])==='ACTIVE';});
 if(!r)throw new Error('참여코드 또는 개인키를 확인하세요.');
 return {code:String(r[0]),name:String(r[1])};
}
function safe_(v){var s=String(v);return /^[\s]*[=+@-]/.test(s)?"'"+s:s;}
function history_(){return rows_(book_().getSheetByName(NAMES_.history));}
function latest_(all,code,caseId){return all.filter(function(r){return String(r[2])===code&&String(r[3])===caseId;}).reduce(function(best,r){return !best||Number(r[4])>Number(best[4])?r:best;},null);}
function result_(person,cfg,r){return {person:person,config:cfg,record:r?{revision:Number(r[4]),updated:String(r[1]),fact:String(r[5]),interpretation:String(r[6]),action:String(r[7]),change:String(r[8])}:{revision:0,updated:'',fact:'',interpretation:'',action:'',change:''}};}
function updateCurrent_(person,r){
 if(!r)return;
 var sh=book_().getSheetByName(NAMES_.current),rows=rows_(sh);
 var i=rows.findIndex(function(x){return String(x[0])===person.code&&String(x[2])===String(r[3]);});
 sh.getRange(i<0?sh.getLastRow()+1:i+2,1,1,9).setValues([[person.code,safe_(person.name),r[3],Number(r[4]),r[1],r[5],r[6],r[7],r[8]]]);
}
function rebuildCurrent_(){
 var lock=LockService.getScriptLock();lock.waitLock(15000);
 try{
  var ss=book_(),people=rows_(ss.getSheetByName(NAMES_.people)),map={};
  history_().forEach(function(r){var key=String(r[2])+'|'+String(r[3]);if(!map[key]||Number(r[4])>Number(map[key][4]))map[key]=r;});
  var sh=ss.getSheetByName(NAMES_.current);if(sh.getLastRow()>1)sh.getRange(2,1,sh.getLastRow()-1,9).clearContent();
  Object.keys(map).forEach(function(k){var r=map[k],person=people.find(function(p){return String(p[0])===String(r[2]);});updateCurrent_({code:String(r[2]),name:person?String(person[1]):String(r[2])},r);});SpreadsheetApp.flush();
 }finally{lock.releaseLock();}
 SpreadsheetApp.getUi().alert('수정이력을 기준으로 현재기록을 다시 만들었습니다.');
}
