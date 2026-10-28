// Tift County Schools feedback backend: Google Apps Script bound to a Google Sheet (same pattern as tsshourtracker).
// Setup: Project Settings > Script Properties: DASHBOARD_CODE=1234, GEMINI_API_KEY=<key>. Never put them in this file.
// Deploy > New deployment > Web app, Execute as: Me, Access: Anyone. Paste the /exec URL into config.js (API const).
// Email is STUBBED: no MailApp/GmailApp calls. Confirmation template is in EMAIL_TEMPLATE.md awaiting approval.
const CATS=['sports','education','safety','facilities','transportation','other'],SEVS=['low','medium','high','urgent'];
const COLS=['id','created_at','name','email','phone','issue_type','feedback','has_child','schools','grades','relationship','may_contact','contact_method','category','severity','ai_status','ai_reason','email_status'];
const P=()=>PropertiesService.getScriptProperties();
function sh(){const ss=SpreadsheetApp.getActive();let s=ss.getSheetByName('Submissions');if(!s){s=ss.insertSheet('Submissions');s.appendRow(COLS);}return s;}
function out(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}
function doPost(e){try{const b=JSON.parse(e.postData.contents);if(b.action==='submit')return out(submit(b));if(b.action==='list')return out(list(b));return out({ok:false,error:'Unknown action'});}
 catch(err){return out({ok:false,error:'Server error'});}}
function cl(v,n){return String(v==null?'':v).slice(0,n||5000);}
function submit(b){if(!b.name||!b.email||!b.phone||!b.feedback)return{ok:false,error:'Missing required fields'};
 const lock=LockService.getScriptLock();lock.waitLock(10000);const s=sh(),id=Math.max(0,s.getLastRow()-1)+1;
 const issue=CATS.indexOf(String(b.issue_type).toLowerCase())>=0?String(b.issue_type).toLowerCase():'';
 const ai=classify(issue,cl(b.feedback));const arr=v=>Array.isArray(v)?v.map(x=>cl(x,80)).join('; '):cl(v,300);
 s.appendRow([id,new Date().toISOString(),cl(b.name,200),cl(b.email,200),cl(b.phone,50),issue,cl(b.feedback),cl(b.has_child,5),arr(b.schools),arr(b.grades),cl(b.relationship,50),cl(b.may_contact,5),cl(b.contact_method,30),ai.category||issue,ai.severity||'',ai.status,ai.reason||'','stubbed_not_sent']);
 lock.releaseLock();return{ok:true,id:id};}
function list(b){const want=P().getProperty('DASHBOARD_CODE');if(!want||String(b.code)!==want){Utilities.sleep(800);return{ok:false,error:'Incorrect code'};}
 const s=sh(),v=s.getDataRange().getValues();
 if(b.retry&&P().getProperty('GEMINI_API_KEY')){let n=0;for(let i=1;i<v.length&&n<15;i++)if(v[i][15]!=='done'){const ai=classify(v[i][5],v[i][6]);if(ai.status==='done'){s.getRange(i+1,14,1,4).setValues([[ai.category,ai.severity,'done',ai.reason]]);}n++;}}
 const rows=s.getDataRange().getValues().slice(1).map(r=>Object.fromEntries(COLS.map((c,i)=>[c,r[i]])));
 const rank={urgent:0,high:1,medium:2,low:3};rows.sort((a,b)=>((rank[a.severity]??4)-(rank[b.severity]??4))||(b.id-a.id));return{ok:true,rows:rows};}
function classify(issue,text){const key=P().getProperty('GEMINI_API_KEY');if(!key)return{status:'pending'};
 const prompt='You triage community feedback for Tift County Schools (Georgia). Categories: '+CATS.join(', ')+'. Severities: '+SEVS.join(', ')+' (urgent = immediate safety risk to people). Submitter\'s chosen category: '+(issue||'none')+' — confirm or correct it. Reply JSON only: {"category":"...","severity":"...","reason":"one sentence"}.\n\nFeedback:\n'+text;
 try{const r=UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent',{method:'post',contentType:'application/json',headers:{'x-goog-api-key':key},muteHttpExceptions:true,
  payload:JSON.stringify({contents:[{parts:[{text:prompt}]}],generationConfig:{responseMimeType:'application/json'}})});
  if(r.getResponseCode()!==200)return{status:'pending',reason:'AI error '+r.getResponseCode()};
  const j=JSON.parse(JSON.parse(r.getContentText()).candidates[0].content.parts[0].text);
  return{status:'done',category:CATS.indexOf(j.category)>=0?j.category:(issue||'other'),severity:SEVS.indexOf(j.severity)>=0?j.severity:'medium',reason:cl(j.reason,300)};}
 catch(e){return{status:'pending',reason:'AI unavailable'};}}
