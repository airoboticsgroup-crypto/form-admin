/** 사업단 통합관리 자료 서버. 전체관리자 소유의 별도 프로젝트에서 실행합니다. */
const UA_OWNER='airobotics.group@gmail.com',UA_VIEWER='ai.robot.kw@gmail.com';
const UA_MODULES={course:'교과',program:'비교과',ta:'TA',equipment:'장비',material:'재료비',biztrip:'출장',expert:'전문가',manuscript:'원고료',directory:'명단'};
const UA_TABLES=['accounts','people','deployments','activities','audit'];
function up_(){return PropertiesService.getScriptProperties();}
function setupUnifiedAdmin(){
 const a=Session.getActiveUser().getEmail().trim().toLowerCase(),e=Session.getEffectiveUser().getEmail().trim().toLowerCase();if(a!==UA_OWNER||e!==UA_OWNER)throw Error('전체관리자 편집기에서 실행하세요.');
 const l=LockService.getScriptLock();l.waitLock(30000);try{
  let id=up_().getProperty('UA_DB_ID');const b=id?SpreadsheetApp.openById(id):SpreadsheetApp.create('사업단_통합관리_DB');if(!id)up_().setProperty('UA_DB_ID',b.getId());
  for(const name of UA_TABLES){let s=b.getSheetByName(name);if(!s)s=b.insertSheet(name);if(!s.getLastRow()){s.getRange(1,1,1,4).setValues([['ID','자료JSON','버전','변경일시']]);s.setFrozenRows(1);}}
  if(!up_().getProperty('UA_SECRET'))up_().setProperty('UA_SECRET',Array.from({length:4},()=>Utilities.getUuid().replace(/-/g,'')).join(''));
  if(!all_('deployments').length)seedDeployments_();
  Logger.log('준비 완료: '+b.getUrl());Logger.log('UA_SECRET은 프로젝트 설정의 스크립트 속성에서 확인하여 접속 프로젝트에만 복사하세요.');return b.getUrl();
 }finally{l.releaseLock();}
}
function doGet(){return ContentService.createTextOutput('사업단 통합관리 자료 서버');}
function doPost(e){try{const r=verify_(e),data=dispatch_(r.email,r.action,r.input||{});return output_({ok:true,data});}catch(e){return output_({ok:false,error:e.uaSafe?e.message:'처리하지 못했습니다. 연결 설정·자료 접근권한을 전체관리자가 확인하세요.'});}}
function output_(v){return ContentService.createTextOutput(JSON.stringify(v)).setMimeType(ContentService.MimeType.JSON);}
function fail_(s){const e=Error(s);e.uaSafe=true;throw e;}
function text_(v,max=300){const s=String(v==null?'':v).trim();if(s.length>max)fail_('입력 길이를 확인하세요.');return s;}
function mail_(v){const s=text_(v,180).toLowerCase();if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s))fail_('이메일 형식을 확인하세요.');return s;}
function required_(v,n=200){const s=text_(v,n);if(!s)fail_('필수 항목을 입력하세요.');return s;}
function module_(m){if(!Object.prototype.hasOwnProperty.call(UA_MODULES,m))fail_('업무를 확인하세요.');return m;}
function verify_(e){
 if(!e||!e.postData||e.postData.contents.length>400000)fail_('요청 오류');const env=JSON.parse(e.postData.contents),secret=up_().getProperty('UA_SECRET');if(!secret||secret.length<64||typeof env.message!=='string'||! /^[A-Za-z0-9_-]{43}=$/.test(env.signature||''))fail_('연결 인증 실패');
 const expected=Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(env.message,secret));let diff=0;for(let i=0;i<expected.length;i++)diff|=expected.charCodeAt(i)^env.signature.charCodeAt(i);if(diff)fail_('연결 인증 실패');const r=JSON.parse(env.message);
 if(r.version!==1||!Number.isSafeInteger(r.time)||Math.abs(Date.now()-r.time)>90000||! /^[a-f0-9-]{36}$/.test(r.nonce||'')||typeof r.action!=='string')fail_('만료되거나 잘못된 요청입니다.');r.email=mail_(r.email);
 const lock=LockService.getScriptLock();lock.waitLock(10000);try{const c=CacheService.getScriptCache(),k='UA:'+r.nonce;if(c.get(k))fail_('이미 처리된 요청입니다.');c.put(k,'1',180);}finally{lock.releaseLock();}return r;
}
function sheet_(name){if(!UA_TABLES.includes(name))fail_('자료 종류 오류');const id=up_().getProperty('UA_DB_ID');if(!id)fail_('전체관리자 초기 설치가 필요합니다.');const s=SpreadsheetApp.openById(id).getSheetByName(name);if(!s)fail_('관리 자료 시트를 확인하세요.');return s;}
function all_(table){const s=sheet_(table),n=s.getLastRow();if(n>10001)fail_('자료가 많아 관리자가 분할해야 합니다.');return n<2?[]:s.getRange(2,1,n-1,4).getValues().map(r=>JSON.parse(r[1]));}
function get_(t,id){return all_(t).find(r=>r.id===id)||null;}
function encodeRow_(r){const s=JSON.stringify(r);if(s.length>45000)fail_('한 항목의 자료가 너무 깁니다.');return [/^[\s]*[=+@-]/.test(r.id)?"'"+r.id:r.id,s,r.revision,r.updatedAt];}
function put_(t,r){const s=sheet_(t),rows=all_(t),i=rows.findIndex(x=>x.id===r.id);const row=i<0?rows.length+2:i+2;if(row>s.getMaxRows())s.insertRowsAfter(s.getMaxRows(),row-s.getMaxRows());s.getRange(row,1,1,4).setValues([encodeRow_(r)]);return r;}
function locked_(fn){const l=LockService.getScriptLock();l.waitLock(30000);try{return fn();}finally{l.releaseLock();}}
function audit_(actor,action,id,detail){const r={id:Utilities.getUuid(),actor,action,target:id,detail:detail||'',revision:1,updatedAt:new Date().toISOString()};const s=sheet_('audit');s.appendRow(encodeRow_(r));}
function base_(input,old){if(Number(input.revision||0)!==Number(old&&old.revision||0))fail_('다른 변경이 있습니다. 새로고침 후 다시 저장하세요.');return {id:old?old.id:(text_(input.id,180)||Utilities.getUuid()),revision:Number(old&&old.revision||0)+1,updatedAt:new Date().toISOString()};}
function account_(email){
 if(email===UA_OWNER)return {id:email,email,name:'전체관리자',role:'owner',active:true,permissions:Object.fromEntries(Object.keys(UA_MODULES).map(m=>[m,'edit'])),download:true,sign:true,remove:true,allowedKeys:[],revision:0};
 const a=get_('accounts',email);if(email===UA_VIEWER)return {...(a||{}),id:email,email,name:a&&a.name||'조회 담당자',role:'viewer',active:!a||a.active!==false,permissions:Object.fromEntries(Object.keys(UA_MODULES).map(m=>[m,a&&a.permissions&&a.permissions[m]==='none'?'none':'view'])),download:!!(a&&a.download),sign:false,remove:false,allowedKeys:keys_(a&&a.allowedKeys||[]),revision:a&&a.revision||0};
 if(!a||!a.active)fail_('등록되지 않았거나 비활성화된 계정입니다.');return {...a,role:'staff'};
}
function owner_(a){if(a.role!=='owner')fail_('전체관리자만 변경할 수 있습니다.');}
function allow_(a,m,write=false,key){module_(m);const p=a.permissions[m]||'none';if(p==='none'||write&&p!=='edit')fail_('해당 업무 권한이 없습니다.');if(key&&a.allowedKeys&&a.allowedKeys.length&&!a.allowedKeys.includes(key))fail_('배정된 자료만 접근할 수 있습니다.');}
function adminRead_(a){if(!['owner','viewer'].includes(a.role))fail_('관리 설정 조회 권한이 없습니다.');}
function saveAccount_(a,p){owner_(a);const email=mail_(p.email);if(email===UA_OWNER)fail_('전체관리자 계정은 변경할 수 없습니다.');const old=get_('accounts',email),permissions={};for(const m of Object.keys(UA_MODULES)){const value=p.permissions&&p.permissions[m]||'none';if(!['none','view','edit'].includes(value))fail_('권한을 확인하세요.');permissions[m]=email===UA_VIEWER?(value==='none'?'none':'view'):value;}
 const r={...base_({...p,id:email},old),email,name:required_(p.name,80),staffId:text_(p.staffId,60),active:p.active===true,permissions,download:p.download===true,sign:email!==UA_VIEWER&&p.sign===true,remove:email!==UA_VIEWER&&p.remove===true,allowedKeys:keys_(p.allowedKeys)};put_('accounts',r);audit_(a.email,'계정·권한 저장',email,JSON.stringify({active:r.active,permissions:r.permissions,download:r.download}));return r;}
function keys_(v){const a=Array.isArray(v)?v:String(v||'').split(/[\n,]/);return [...new Set(a.map(x=>text_(x,180)).filter(Boolean))].slice(0,200);}
function normalizePerson_(p,old){
 const kind=text_(p.kind,20);if(!['student','professor','ta','staff'].includes(kind))fail_('명단 구분을 확인하세요.');const personNo=required_(p.personNo,60),name=required_(p.name,80),emails=keys_(p.emails).map(mail_);if(!emails.length)fail_('이메일을 하나 이상 입력하세요.');
 const term=kind==='student'?required_(p.term,30):text_(p.term,30);if(term&&!/^20\d{2}-(1|2|여름|겨울)$/.test(term))fail_('학기는 2026-1, 2026-2 형식으로 입력하세요.');
 if(old&&(old.kind!==kind||old.personNo!==personNo))fail_('기존 명단의 구분·학번·사번은 변경할 수 없습니다. 새 항목으로 등록하세요.');
 const memberships=old&&old.memberships?JSON.parse(JSON.stringify(old.memberships)):[];if(kind==='student'){const v={term,active:p.active!==false},i=memberships.findIndex(x=>x.term===term);if(i<0)memberships.push(v);else memberships[i]=v;}
 return {...base_(p,old),kind,personNo,name,department:text_(p.department,150),emails,term,active:kind==='student'?memberships.some(x=>x.active):p.active!==false,memberships,note:text_(p.note,500)};
}
function previewPeople_(a,p){allow_(a,'directory',true);if(!Array.isArray(p.rows)||!p.rows.length||p.rows.length>250)fail_('한 번에 1~250명을 업로드하세요.');const existing=all_('people'),seen=new Set();const rows=p.rows.map((x,i)=>{const key=x.kind+':'+x.personNo;if(seen.has(key))fail_((i+1)+'행: 중복 학번·사번입니다.');seen.add(key);const old=existing.find(r=>r.kind===x.kind&&r.personNo===String(x.personNo||'').trim()),r=normalizePerson_({...x,id:old&&old.id,revision:old&&old.revision||0},old);return {...r,expectedRevision:old&&old.revision||0,change:old?'변경':'추가'};});return {rows,add:rows.filter(x=>x.change==='추가').length,update:rows.filter(x=>x.change==='변경').length};}
function savePerson_(a,p){allow_(a,'directory',true);const old=p.id?get_('people',p.id):null;if(p.id&&!old)fail_('기존 명단을 찾을 수 없습니다.');const r=normalizePerson_(p,old);if(all_('people').some(x=>x.id!==r.id&&x.kind===r.kind&&x.personNo===r.personNo))fail_('이미 등록된 학번·사번입니다.');put_('people',r);audit_(a.email,'명단 저장',r.id,r.kind+' / '+r.personNo);return r;}
function applyPeople_(a,p){allow_(a,'directory',true);if(!Array.isArray(p.rows)||!p.rows.length||p.rows.length>250)fail_('업로드 범위를 확인하세요.');const list=all_('people'),seen=new Set(),next=p.rows.map(x=>{const old=list.find(v=>v.kind===x.kind&&v.personNo===x.personNo),k=x.kind+':'+x.personNo;if(seen.has(k))fail_('업로드 중복');seen.add(k);return normalizePerson_({...x,id:old?old.id:'',revision:x.expectedRevision},old);});for(const r of next){const i=list.findIndex(x=>x.id===r.id);if(i<0)list.push(r);else list[i]=r;}const s=sheet_('people');if(list.length+1>s.getMaxRows())s.insertRowsAfter(s.getMaxRows(),list.length+1-s.getMaxRows());s.getRange(2,1,list.length,4).setValues(list.map(encodeRow_));audit_(a.email,'명단 일괄 반영','people',next.length+'명');return {saved:next.length};}
function cleanDeploy_(p,old){const m=module_(p.module),kind=text_(p.intakeType,20)||'none';if(!['none','course','program'].includes(kind)||(kind!=='none'&&kind!==m))fail_('조회 연결 유형을 확인하세요.');const url=text_(p.url,500);if(url&&!/^https:\/\/(script\.google\.com\/macros\/s\/[\w-]+\/exec(?:\?page=admin)?|airoboticsgroup-crypto\.github\.io\/form\/[^\s]*)$/.test(url))fail_('허용된 Apps Script 또는 사업단 안내 주소를 입력하세요.');for(const k of ['sheetId','folderId'])if(p[k]&&!/^[\w-]{15,150}$/.test(p[k]))fail_('시트·폴더 ID를 확인하세요.');return {...base_(p,old),module:m,name:required_(p.name,100),purpose:text_(p.purpose,80),url,version:text_(p.version,40),sheetId:text_(p.sheetId,150),sheetName:text_(p.sheetName,100)||'접수목록',folderId:text_(p.folderId,150),intakeType:kind,active:p.active===true};}
function deployList_(a){adminRead_(a);return all_('deployments').map(x=>a.role==='owner'?x:{...x,sheetId:'',folderId:''});}
function source_(m){const rows=all_('deployments').filter(x=>x.module===m&&x.active&&x.intakeType===m&&x.sheetId&&x.folderId);if(rows.length!==1)fail_('업무별 자료 연결을 하나만 활성화하세요.');return rows[0];}
function applications_(a,p){allow_(a,p.module);if(!['course','program'].includes(p.module))return {connected:false,rows:[],note:'이 업무의 접수자료 연결은 별도 연동이 필요합니다.'};let c;try{c=source_(p.module);}catch(e){if(e.uaSafe)return {connected:false,rows:[],note:e.message};throw e;}const s=SpreadsheetApp.openById(c.sheetId).getSheetByName(c.sheetName);if(!s)fail_('연결된 접수 시트가 없습니다.');const n=s.getLastRow();if(n>20001)fail_('조회 범위를 나누어야 합니다.');const rows=n<2?[]:s.getRange(2,1,n-1,29).getDisplayValues().map(r=>({receipt:r[0],date:r[1],title:r[2],name:r[6],total:r[16],status:r[20],id:r[28]})).filter(r=>!a.allowedKeys.length||a.allowedKeys.includes(r.id)).reverse();return {connected:true,rows,note:'신청내용 조회 연결 완료. 서명·기안 등 처리는 기존 업무 화면에서 진행합니다.'};}
function record_(a,p){allow_(a,p.module,false,p.id);const c=source_(p.module);if(!/^[\w-]{15,150}$/.test(p.id||''))fail_('자료 ID 오류');const rows=applications_(a,p).rows;if(!rows.some(r=>r.id===p.id))fail_('접수목록에 없는 자료입니다.');const f=DriveApp.getFileById(p.id);if(!inFolder_(f,c.folderId))fail_('자료 위치 오류');return {c,data:JSON.parse(f.getBlob().getDataAsString('UTF-8'))};}
function inFolder_(file,id){const it=file.getParents();while(it.hasNext())if(it.next().getId()===id)return true;return false;}
function review_(a,p){const {data}=record_(a,p),payload=JSON.parse(JSON.stringify(data.payload));delete payload.signature;delete payload.requestKey;return {receipt:data.receipt,payload,workflow:data.workflow||{},signed:!!data.finalId};}
function pdf_(a,p){if(!a.download)fail_('다운로드 권한이 없습니다.');const {c,data}=record_(a,p),id=data.finalId||data.draftId;if(!id)fail_('PDF가 없습니다.');const f=DriveApp.getFileById(id);if(!inFolder_(f,c.folderId))fail_('PDF 위치 오류');const b=f.getBlob(),bytes=b.getBytes();if(bytes.length>15000000)fail_('파일 크기를 확인하세요.');audit_(a.email,'PDF 다운로드',p.id,p.module);return {filename:data.receipt+'.pdf',base64:Utilities.base64Encode(bytes),mime:'application/pdf'};}
function saveActivity_(a,p){allow_(a,'ta',true,p.id||p.fileId);const old=p.id?get_('activities',p.id):null;if(p.id&&!old)fail_('활동일지 등록정보가 없습니다.');const fileId=required_(p.fileId,150);if(old&&old.fileId!==fileId)allow_(a,'ta',true,fileId);if(!/^[\w-]{15,150}$/.test(fileId))fail_('파일 ID 오류');if(all_('activities').some(x=>x.id!==p.id&&x.fileId===fileId))fail_('이미 등록된 활동일지 파일입니다.');const month=text_(p.month,7);if(!/^20\d{2}-(0[1-9]|1[0-2])$/.test(month))fail_('활동월을 확인하세요.');const r={...base_(p,old),name:required_(p.name,80),personNo:required_(p.personNo,60),course:required_(p.course,150),taType:p.taType==='practice'?'practice':'education',month,fileId,note:text_(p.note,500),shareStatus:old&&old.fileId===fileId?old.shareStatus||'미적용':'미적용'};put_('activities',r);audit_(a.email,'TA 활동일지 연결 저장',r.id,r.month);return r;}
function syncActivity_(a,p){owner_(a);const r=get_('activities',p.id);if(!r)fail_('활동일지를 선택하세요.');const f=DriveApp.getFileById(r.fileId);if(f.isTrashed())fail_('휴지통 파일입니다.');const owner=f.getOwner();if(!owner||owner.getEmail().toLowerCase()===UA_VIEWER)fail_('파일 소유권을 확인하세요.');if(owner.getEmail().toLowerCase()!==UA_OWNER)f.addEditor(UA_OWNER);f.removeEditor(UA_VIEWER);f.addViewer(UA_VIEWER);const access=f.getAccess(UA_VIEWER);if(access!==DriveApp.Permission.VIEW)fail_('상위 폴더·그룹 권한 때문에 조회 전용 적용을 확인할 수 없습니다. 공유 설정을 점검하세요.');r.shareStatus='적용 완료';r.updatedAt=new Date().toISOString();r.revision++;put_('activities',r);audit_(a.email,'TA 공유권한 적용',r.id,'전체관리자 편집 / 지정계정 조회');return {name:f.getName(),status:r.shareStatus};}
function dispatch_(email,action,p){
 const a=account_(mail_(email));if(!a.active)fail_('비활성화된 계정입니다.');const reads={
  'bootstrap':()=>({account:a,modules:UA_MODULES,version:'통합관리 1.0',counts:{people:['view','edit'].includes(a.permissions.directory)?all_('people').length:null,accounts:['owner','viewer'].includes(a.role)?all_('accounts').filter(r=>r.email!==UA_VIEWER&&r.email!==UA_OWNER).length+2:null,activities:['view','edit'].includes(a.permissions.ta)?all_('activities').filter(r=>!a.allowedKeys.length||a.allowedKeys.includes(r.id)).length:null}}),
  'accounts.list':()=>{adminRead_(a);return [account_(UA_OWNER),account_(UA_VIEWER),...all_('accounts').filter(x=>x.email!==UA_VIEWER)];},
  'people.list':()=>{allow_(a,'directory');return all_('people');},
  'people.preview':()=>previewPeople_(a,p),
  'people.export':()=>{allow_(a,'directory');if(!a.download)fail_('다운로드 권한이 없습니다.');audit_(a.email,'명단 다운로드','people');return all_('people');},
  'deployments.list':()=>deployList_(a),
  'activities.inspect':()=>{owner_(a);const r=get_('activities',p.id);if(!r)fail_('활동일지를 선택하세요.');const f=DriveApp.getFileById(r.fileId);return {name:f.getName(),fileId:r.fileId};},
  'activities.list':()=>{allow_(a,'ta');return all_('activities').filter(r=>!a.allowedKeys.length||a.allowedKeys.includes(r.id));},
  'applications.deletePreview':()=>previewCleanup_(a,p),'applications.list':()=>applications_(a,p),'applications.review':()=>review_(a,p),'applications.pdf':()=>pdf_(a,p),
  'audit.list':()=>all_('audit').filter(x=>['owner','viewer'].includes(a.role)||x.actor===a.email).slice(-300).reverse()
 };if(Object.prototype.hasOwnProperty.call(reads,action))return reads[action]();
 const writeActions=['accounts.save','people.save','people.import','deployments.save','activities.save','activities.sync','applications.delete'];if(!writeActions.includes(action))fail_('지원하지 않는 작업입니다.');return locked_(()=>{const fresh=account_(email);if(!fresh.active)fail_('계정이 비활성화되었습니다.');return writesWithFresh_(fresh,action,p);});
}
function writesWithFresh_(a,action,p){
 switch(action){case 'applications.delete':return deleteCleanup_(a,p);case 'accounts.save':return saveAccount_(a,p);case 'people.save':return savePerson_(a,p);case 'people.import':return applyPeople_(a,p);case 'activities.save':return saveActivity_(a,p);case 'activities.sync':return syncActivity_(a,p);case 'deployments.save':owner_(a);const old=p.id?get_('deployments',p.id):null;if(p.id&&!old)fail_('기존 연결정보가 없습니다.');const r=cleanDeploy_(p,old);if(r.active&&r.intakeType!=='none'&&all_('deployments').some(x=>x.id!==r.id&&x.module===r.module&&x.active&&x.intakeType===r.intakeType))fail_('동일 업무의 자료 연결은 하나만 활성화하세요.');put_('deployments',r);audit_(a.email,'업무 연결 저장',r.id,r.name);return r;default:fail_('작업 오류');}
}

function seedDeployments_(){
 const courseApply='AKfycbxKTHeXOv4GSCmXvhBKmohPLNRDb47BrrX7Rxm00mBIDhPRryGaV776HN2S0uxqUSmx3w',courseAdmin='AKfycbwYo3AlZUvzpH0Jj62W69MUUGQprugpXUZFFOL1MPH2h-A1mukw4vrtqq4mUb8R-J0v2w',programApply='AKfycbzyVi76Sjfmb2csn3W-w7hryLDigyOTUa_q5nV46RSo_ho_FuY59hVEhrtGzFXhB1S9',programAdmin='AKfycbwx283NtY8ylFfr3wJqZ2_GPCvIOqSL7t3vJJsiE9LdB4ENOrspV1JwzT_METggjLg';
 for(const [id,module,name,purpose,deployment]of [['course-apply','course','[양식1] 신청용','신청용',courseApply],['course-admin','course','[양식1] 관리자용','관리자용',courseAdmin],['program-apply','program','[양식2] 신청용','신청용',programApply],['program-admin','program','[양식2] 관리자용','관리자용',programAdmin]])put_('deployments',{id,module,name,purpose,url:'https://script.google.com/macros/s/'+deployment+'/exec'+(purpose==='관리자용'?'?page=admin':''),version:'',sheetId:'',sheetName:'접수목록',folderId:'',intakeType:'none',active:true,revision:1,updatedAt:new Date().toISOString()});
}
