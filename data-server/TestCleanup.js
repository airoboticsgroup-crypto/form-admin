/** 전체관리자 전용 테스트 접수 정리. 영구 삭제하지 않고 관련 파일을 휴지통으로 보냅니다. */
function cleanupKey_(token){if(!/^[a-f0-9-]{36}$/i.test(token||''))fail_('삭제 확인이 만료되었습니다. 다시 선택하세요.');return 'UA_DELETE:'+token;}
function cleanupFingerprint_(value){return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(JSON.stringify(value),up_().getProperty('UA_SECRET')));}
function cleanupSelection_(a,p){owner_(a);if(!a.remove)fail_('삭제 권한이 없습니다.');if(!['course','program'].includes(p.module))fail_('교과·비교과 테스트 접수만 정리할 수 있습니다.');if(!Array.isArray(p.ids)||p.ids.length<1||p.ids.length>10||p.ids.some(id=>typeof id!=='string'||! /^[\w-]{15,150}$/.test(id))||new Set(p.ids).size!==p.ids.length)fail_('서로 다른 접수를 1~10건 선택하세요.');return p.ids;}
function cleanupPlan_(a,module,id){
 owner_(a);const c=source_(module),book=SpreadsheetApp.openById(c.sheetId),sheet=book.getSheetByName(c.sheetName);if(!sheet)fail_('접수 시트가 없습니다.');const n=sheet.getLastRow();if(n>20001)fail_('접수자료가 많아 조회 범위를 나누어야 합니다.');
 const rows=n<2?[]:sheet.getRange(2,1,n-1,29).getDisplayValues(),matches=rows.map((r,i)=>({r,row:i+2})).filter(x=>x.r[28]===id);if(matches.length!==1)fail_('접수목록에 없거나 중복된 자료입니다. 새로고침하세요.');const {r,row}=matches[0],file=DriveApp.getFileById(id);
 if(file.isTrashed()||!inFolder_(file,c.folderId))fail_('접수자료 위치를 확인하세요.');const text=file.getBlob().getDataAsString('UTF-8');let d;try{d=JSON.parse(text);}catch(e){fail_('접수자료 형식을 확인하세요.');}
 const prefix=module==='course'?'CP':'EP';if(rows.filter(x=>x[0]===d.receipt).length!==1)fail_('같은 접수번호의 자료가 여러 개입니다. 개별 확인이 필요합니다.');if(!new RegExp('^'+prefix+'-20\\d{6}-[A-F0-9]{16}$').test(d.receipt||'')||r[0]!==d.receipt||!d.payload||r[26]!==d.payload.requestKey||r[27]!==d.hash||file.getName()!==d.payload.requestKey+'.json')fail_('접수번호와 원본자료가 일치하지 않습니다.');
 if(d.pdfJob&&['queued','working'].includes(d.pdfJob.state))fail_('PDF 생성 중인 접수는 완료 후 정리하세요: '+d.receipt);
 let queue=null;const q=book.getSheetByName('PDF생성대기열');if(q&&q.getLastRow()>1){const headers=q.getRange(1,1,1,2).getDisplayValues()[0];if(headers[0]!=='자료파일 ID'||headers[1]!=='상태')fail_('PDF 대기열 형식을 확인하세요.');const items=q.getRange(2,1,q.getLastRow()-1,8).getValues(),found=items.map((v,i)=>({v,row:i+2})).filter(x=>x.v[0]===id);if(found.length>1)fail_('PDF 대기열의 중복 자료를 확인하세요.');if(found.length){queue=found[0];if(!['done','error','cancelled'].includes(queue.v[1]))fail_('PDF 생성 중인 접수는 완료 후 정리하세요: '+d.receipt);}}
 const references=[d.draftId,d.finalId,...(d.pdfHistory||[]).flatMap(x=>[x.draftId,x.finalId]),...(d.businessHistory||[]).map(x=>x.pdfId)].filter(Boolean),pdfIds=[...new Set(references)];if(pdfIds.length>30)fail_('관련 PDF가 많습니다. 접수 한 건씩 확인하세요.');
 if(rows.some(x=>x[28]!==id&&[x[19],x[23]].some(url=>url&&pdfIds.some(pdfId=>url==='https://drive.google.com/file/d/'+pdfId+'/view'))))fail_('다른 접수가 함께 참조하는 PDF는 정리할 수 없습니다.');
 const pdfs=pdfIds.map(pdfId=>{if(!/^[\w-]{15,150}$/.test(pdfId)||pdfId===id)fail_('PDF 자료 ID를 확인하세요.');const f=DriveApp.getFileById(pdfId);if(!inFolder_(f,c.folderId)||f.getMimeType()!=='application/pdf'||!f.getName().startsWith(d.receipt+'_'))fail_('다른 접수의 파일은 삭제할 수 없습니다.');return {id:pdfId,name:f.getName(),trashed:f.isTrashed()};});
 for(const url of [r[19],r[23]].filter(Boolean)){const m=url.match(/^https:\/\/drive\.google\.com\/file\/d\/([\w-]+)\/view$/);if(!m||!pdfIds.includes(m[1]))fail_('목록의 PDF와 원본자료가 일치하지 않습니다.');}
 const manifest={module,id,receipt:d.receipt,sheetId:c.sheetId,sheetName:c.sheetName,folderId:c.folderId,rowValues:r,record:text,pdfs,queue:queue&&queue.v};
 return {id,module,receipt:d.receipt,title:r[2],date:r[1],name:r[6],fileCount:pdfs.filter(x=>!x.trashed).length+1,fingerprint:cleanupFingerprint_(manifest),pdfs,queue,sheet,q,row};
}
function previewCleanup_(a,p){
 const ids=cleanupSelection_(a,p),plans=ids.map(id=>cleanupPlan_(a,p.module,id));if(plans.reduce((n,x)=>n+x.fileCount,0)>80)fail_('관련 파일이 많습니다. 선택 건수를 줄이세요.');
 const token=Utilities.getUuid(),items=plans.map(x=>({id:x.id,receipt:x.receipt,fingerprint:x.fingerprint}));CacheService.getScriptCache().put(cleanupKey_(token),JSON.stringify({email:a.email,module:p.module,items}),300);
 return {token,rows:plans.map(x=>({id:x.id,receipt:x.receipt,title:x.title,date:x.date,name:x.name,fileCount:x.fileCount})),note:'테스트 접수인지 확인하세요. 접수목록의 해당 행을 삭제하고 관련 PDF·자료파일을 휴지통으로 보냅니다.'};
}
function deleteCleanup_(a,p){
 owner_(a);if(!a.remove)fail_('삭제 권한이 없습니다.');if(p.confirmation!=='테스트 삭제'||p.testOnly!==true)fail_('테스트 자료 확인과 확인 문구가 필요합니다.');
 const key=cleanupKey_(p.token),cache=CacheService.getScriptCache(),raw=cache.get(key);if(!raw)fail_('삭제 확인이 만료되었습니다. 다시 선택하세요.');const saved=JSON.parse(raw);if(saved.email!==a.email)fail_('삭제 확인 계정이 다릅니다.');
 cleanupSelection_(a,{module:saved.module,ids:saved.items.map(x=>x.id)});const plans=saved.items.map(x=>{const plan=cleanupPlan_(a,saved.module,x.id);if(plan.fingerprint!==x.fingerprint)fail_('선택 후 자료가 변경되었습니다. 다시 확인하세요: '+x.receipt);return plan;});
 // 모든 항목을 검증한 다음, 해당 확인 요청은 한 번만 사용합니다.
 cache.remove(key);const results=[];
 for(const plan of plans){
  const moved=[];let queueChanged=false,removed=false;
  try{
   audit_(a.email,'테스트 접수 정리 시작',plan.id,JSON.stringify({module:plan.module,receipt:plan.receipt,pdfIds:plan.pdfs.map(x=>x.id)}));
   if(plan.queue){const v=plan.queue.v.slice();v[1]='cancelled';v[3]=0;v[7]=new Date().toISOString();queueChanged=true;plan.q.getRange(plan.queue.row,1,1,8).setValues([v]);}
   for(const item of [...plan.pdfs.filter(x=>!x.trashed),{id:plan.id}]){const f=DriveApp.getFileById(item.id);moved.push(item.id);f.setTrashed(true);}
   // 앞 항목 삭제로 행 번호가 바뀌므로 매번 자료 ID로 다시 찾습니다.
   const n=plan.sheet.getLastRow(),ids=n<2?[]:plan.sheet.getRange(2,29,n-1,1).getDisplayValues();const i=ids.findIndex(x=>x[0]===plan.id);if(i<0)fail_('접수 행이 변경되었습니다.');plan.sheet.deleteRow(i+2);removed=true;
   try{audit_(a.email,'테스트 접수 정리 완료',plan.id,plan.receipt+' / 파일 '+moved.length+'개 휴지통 이동');}catch(e){results.push({receipt:plan.receipt,deleted:true,warning:'변경 이력의 완료 기록을 확인하세요.'});continue;}
   results.push({receipt:plan.receipt,deleted:true});
  }catch(e){
   // 행 삭제 전 실패하면 이번에 옮긴 파일만 복원합니다. 이미 휴지통에 있던 파일은 건드리지 않습니다.
   const n=plan.sheet.getLastRow(),stillListed=n>1&&plan.sheet.getRange(2,29,n-1,1).getDisplayValues().some(x=>x[0]===plan.id);
   if(!removed&&!stillListed){results.push({receipt:plan.receipt,deleted:true,warning:'목록 삭제는 완료됐습니다. 변경 이력을 확인하세요.'});continue;}
   let rollbackFailed=false;for(const id of moved.reverse())try{DriveApp.getFileById(id).setTrashed(false);}catch(restoreError){rollbackFailed=true;}
   if(queueChanged)try{plan.q.getRange(plan.queue.row,1,1,8).setValues([plan.queue.v]);}catch(restoreError){rollbackFailed=true;}
   const message=rollbackFailed?'일부 파일의 복원을 확인해야 합니다. 접수목록과 Drive 휴지통을 확인하세요.':'정리하지 못했습니다. 접수목록을 유지하고 이동한 파일을 복원했습니다. 다시 확인하세요.';
   try{audit_(a.email,'테스트 접수 정리 실패',plan.id,plan.receipt+' / '+message);}catch(auditError){}
   results.push({receipt:plan.receipt,deleted:false,error:message});
  }
 }
 SpreadsheetApp.flush();return {results,deleted:results.filter(x=>x.deleted).length,failed:results.filter(x=>!x.deleted).length};
}
