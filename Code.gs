const SHEET_ID="1Oas8ItDyfy8lu-JhbOTFz1PS-4fQ25SXNwFux3qC6CE";
const ROOT_FOLDER_ID="1HUpChEpSyYomjn889r3bbaQDimkQbH0V";
const PINS={teacher:"256901",inspector:"256902",admin:"256999"};
const TOKEN_TTL=21600;
const DEFAULT_TERM={id:"2-2569",semester:2,year:2569,start:"2026-11-01",end:"2027-03-31"};

function doPost(e){
 try{
  const p=JSON.parse(e.postData.contents||"{}");
  if(p.action==="login")return out(login(p));
  const role=verify(p.token);if(!role)return out({ok:false,error:"Unauthorized"});
  const routes={
   checkin:()=>allow(role,["teacher","admin"],()=>saveCheckin(p)),
   evaluation:()=>allow(role,["inspector","admin"],()=>saveEvaluation(p)),
   checkinForArea:()=>allow(role,["inspector","admin"],()=>findCheckin(p)),
   dashboard:()=>allow(role,["inspector","admin"],()=>dashboard(p)),
   adminDaily:()=>allow(role,["admin"],()=>adminDaily(p)),
   periodReport:()=>allow(role,["admin"],()=>periodReport(p)),
   listTerms:()=>allow(role,["admin"],()=>listTerms()),
   saveTerm:()=>allow(role,["admin"],()=>saveTerm(p))
  };
  return out(routes[p.action]?routes[p.action]():{ok:false,error:"Unknown action"});
 }catch(err){return out({ok:false,error:String(err&&err.message||err)})}
}
function allow(role,roles,fn){return roles.indexOf(role)>=0?fn():{ok:false,error:"Permission denied"}}
function login(p){if(!PINS[p.role]||String(p.pin)!==PINS[p.role])return {ok:false};const token=Utilities.getUuid()+Utilities.getUuid();CacheService.getScriptCache().put("tok_"+token,p.role,TOKEN_TTL);return {ok:true,role:p.role,token:token}}
function verify(t){return t?CacheService.getScriptCache().get("tok_"+t):null}
function out(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON)}
function ss(){return SpreadsheetApp.openById(SHEET_ID)}
function getSheet(name,headers){let s=ss().getSheetByName(name);if(!s)s=ss().insertSheet(name);if(s.getLastRow()===0)s.appendRow(headers);else{const h=s.getRange(1,1,1,s.getLastColumn()).getValues()[0];headers.forEach(x=>{if(h.indexOf(x)<0){s.getRange(1,s.getLastColumn()+1).setValue(x);h.push(x)}})}return s}
function objects(name,headers){const s=getSheet(name,headers),v=s.getDataRange().getValues(),h=v.shift();return v.map(r=>Object.fromEntries(h.map((k,i)=>[k,r[i]])))}
function safe(s){return String(s||"").replace(/[\\/:*?"<>|#%{}]/g,"-").trim()}
function sub(parent,n){const it=parent.getFoldersByName(n);return it.hasNext()?it.next():parent.createFolder(n)}
function uploadImages(p){if(!p.images||!p.images.length)return[];const root=DriveApp.getFolderById(ROOT_FOLDER_ID),tf=sub(root,"CheckIn_"+safe(p.termId||DEFAULT_TERM.id)),m=sub(tf,String(p.date).slice(0,7)),d=sub(m,safe(p.date)),a=sub(d,"พื้นที่_"+safe(p.area)),urls=[];p.images.forEach((im,i)=>{const blob=Utilities.newBlob(Utilities.base64Decode(im.data),im.mime||"image/jpeg",safe(p.date)+"_พื้นที่"+safe(p.area)+"_"+safe(p.teacher)+"_"+(i+1)+".jpg");urls.push(a.createFile(blob).getUrl())});return urls}
const CH=["Timestamp","Date","Area","Teacher","Room","Reason","Note","PhotoURLs","TermID"];
const EH=["Timestamp","Date","Area","Inspector","Scores","Total","Percent","Note","TermID"];
const TH=["TermID","Semester","AcademicYear","StartDate","EndDate","Active"];
function termOf(o){return String(o.TermID||DEFAULT_TERM.id)}
function saveCheckin(p){
 const s=getSheet("CheckIn",CH),values=s.getDataRange().getValues(),headers=values[0]||CH;
 const idx={date:headers.indexOf("Date"),area:headers.indexOf("Area"),teacher:headers.indexOf("Teacher"),term:headers.indexOf("TermID")};
 const termId=String(p.termId||DEFAULT_TERM.id);
 let replaceRow=0;
 for(let i=1;i<values.length;i++){
  const r=values[i],rowTerm=idx.term>=0&&r[idx.term]?String(r[idx.term]):DEFAULT_TERM.id;
  if(rowTerm===termId&&dateStr(r[idx.date])===String(p.date)&&Number(r[idx.area])===Number(p.area)&&String(r[idx.teacher])===String(p.teacher))replaceRow=i+1;
 }
 const u=uploadImages(p);
 const row=[new Date(),p.date,p.area,p.teacher,p.room,p.reason||"",p.note||"",u.join("\n"),termId];
 if(replaceRow){s.getRange(replaceRow,1,1,row.length).setValues([row]);return {ok:true,replaced:true,photoUrls:u};}
 s.appendRow(row);return {ok:true,replaced:false,photoUrls:u};
}
function saveEvaluation(p){const s=getSheet("Evaluation",EH);s.appendRow([new Date(),p.date,p.area,p.inspector,(p.scores||[]).join(","),p.sum,p.pct,p.note||"",p.termId||DEFAULT_TERM.id]);return {ok:true}}
function checkRows(p){return objects("CheckIn",CH).filter(r=>termOf(r)===String(p.termId||DEFAULT_TERM.id)).map(r=>({date:dateStr(r.Date),area:r.Area,teacher:r.Teacher,room:r.Room,reason:r.Reason,note:r.Note,photoUrls:String(r.PhotoURLs||"").split("\n").filter(Boolean)}))}
function evalRows(p){return objects("Evaluation",EH).filter(r=>termOf(r)===String(p.termId||DEFAULT_TERM.id)).map(r=>({date:dateStr(r.Date),area:r.Area,inspector:r.Inspector,pct:r.Percent,note:r.Note}))}
function dateStr(v){if(v instanceof Date)return Utilities.formatDate(v,Session.getScriptTimeZone(),"yyyy-MM-dd");return String(v||"").slice(0,10)}
function findCheckin(p){const v=checkRows(p).filter(x=>x.date===String(p.date)&&Number(x.area)===Number(p.area));return {ok:true,data:v.length?v[v.length-1]:null}}
function dashboard(p){return {ok:true,checkins:checkRows(p).filter(x=>x.date===String(p.date)),evaluations:evalRows(p).filter(x=>x.date===String(p.date))}}
function adminDaily(p){return dashboard(p)}
function periodReport(p){return {ok:true,checkins:checkRows(p).filter(x=>x.date>=p.start&&x.date<=p.end),evaluations:evalRows(p).filter(x=>x.date>=p.start&&x.date<=p.end)}}
function termSheet(){return getSheet("Terms",TH)}
function ensureTerm(){const s=termSheet(),v=objects("Terms",TH);if(!v.length)s.appendRow([DEFAULT_TERM.id,DEFAULT_TERM.semester,DEFAULT_TERM.year,DEFAULT_TERM.start,DEFAULT_TERM.end,true])}
function listTerms(){ensureTerm();const v=objects("Terms",TH);return {ok:true,terms:v.map(r=>({id:String(r.TermID),semester:Number(r.Semester),year:Number(r.AcademicYear),start:dateStr(r.StartDate),end:dateStr(r.EndDate),active:r.Active===true||String(r.Active).toLowerCase()==="true"}))}}
function saveTerm(p){const s=termSheet(),v=objects("Terms",TH);if(v.some(r=>String(r.TermID)===String(p.id)))return {ok:false,error:"มีภาคเรียนนี้อยู่แล้ว"};if(s.getLastRow()>1)s.getRange(2,6,s.getLastRow()-1,1).setValue(false);s.appendRow([p.id,p.semester,p.year,p.start,p.end,true]);return {ok:true}}
