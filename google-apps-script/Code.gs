/** Project Management student forms backend.
 * Set DRIVE_FOLDER_ID to the ID ONLY of your top-level Drive folder, then deploy as a Web App.
 * The script creates Submissions/Application Data and one readable folder per project/team.
 */
const DRIVE_FOLDER_ID = '1B8YstCdeBnwC5raaTvYwbDALb2B-6m_w';
const INDEX_FILE = 'project-management-form-index.json';

function doPost(e){try{const req=JSON.parse(e.postData.contents);if(req.action!=='save')throw new Error('Unsupported action');return json_({ok:true,...saveRecord_(req.formType,req.mode,req.data||{})});}catch(err){return json_({ok:false,error:String(err.message||err)});}}
function doGet(e){try{if(e.parameter.action!=='load')throw new Error('Unsupported action');return json_({ok:true,data:loadRecord_(e.parameter.formType,e.parameter.recordId)});}catch(err){return json_({ok:false,error:String(err.message||err)});}}
function root_(){return DriveApp.getFolderById(DRIVE_FOLDER_ID);}
function getOrCreateFolder_(parent,name){const it=parent.getFoldersByName(name);return it.hasNext()?it.next():parent.createFolder(name);}
function submissions_(){return getOrCreateFolder_(root_(),'Submissions');}
function dataFolder_(){return getOrCreateFolder_(submissions_(),'Application Data');}
function index_(){const f=dataFolder_(),it=f.getFilesByName(INDEX_FILE);if(!it.hasNext())return {};return JSON.parse(it.next().getBlob().getDataAsString()||'{}');}
function writeIndex_(idx){const f=dataFolder_(),it=f.getFilesByName(INDEX_FILE),s=JSON.stringify(idx,null,2);if(it.hasNext())it.next().setContent(s);else f.createFile(INDEX_FILE,s,MimeType.PLAIN_TEXT);}
function clean_(s){return String(s||'Untitled').replace(/[\\/:*?"<>|#%{}~]/g,'-').replace(/\s+/g,' ').trim().substring(0,100)||'Untitled';}
function teamFolder_(data){return getOrCreateFolder_(submissions_(),clean_(data.projectTitle)+' - '+clean_(data.studentTeam));}

function saveRecord_(formType,mode,data){
  const df=dataFolder_(),idx=index_(); let id=data.recordId||Utilities.getUuid(); const now=new Date().toISOString();
  const rec={recordId:id,formType,mode,updatedAt:now,data:{...data,recordId:id}};
  let entry=idx[id]||{}; if(typeof entry==='string') entry={jsonFileId:entry};
  let jf=null; if(entry.jsonFileId){try{jf=DriveApp.getFileById(entry.jsonFileId);jf.setContent(JSON.stringify(rec,null,2));}catch(e){jf=null;}}
  if(!jf){jf=df.createFile(formType+'_'+clean_(data.studentTeam)+'_'+id+'.json',JSON.stringify(rec,null,2),MimeType.PLAIN_TEXT);entry.jsonFileId=jf.getId();}
  if(formType==='project-plan') entry.readableFileId=upsertPlanDoc_(entry.readableFileId,data,mode,now);
  if(formType==='project-schedule') entry.readableFileId=upsertScheduleSheet_(entry.readableFileId,data,mode,now);
  idx[id]=entry; writeIndex_(idx);
  return {recordId:id,jsonFileId:entry.jsonFileId,readableFileId:entry.readableFileId,updatedAt:now};
}
function loadRecord_(formType,id){if(!id)throw new Error('Missing record ID');const idx=index_(),entry=idx[id];if(!entry)throw new Error('Record not found');const fileId=typeof entry==='string'?entry:entry.jsonFileId;const rec=JSON.parse(DriveApp.getFileById(fileId).getBlob().getDataAsString());if(rec.formType!==formType)throw new Error('Form type mismatch');return rec.data;}

function upsertPlanDoc_(fileId,d,mode,now){
  const folder=teamFolder_(d),name=clean_(d.projectTitle)+' - '+clean_(d.studentTeam)+' - Project Plan'; let doc;
  try{doc=fileId?DocumentApp.openById(fileId):null;}catch(e){doc=null;}
  if(!doc){doc=DocumentApp.create(name);moveToFolder_(DriveApp.getFileById(doc.getId()),folder);} else doc.setName(name);
  const b=doc.getBody();b.clear(); b.appendParagraph('PROJECT PLAN AND SCOPE').setHeading(DocumentApp.ParagraphHeading.TITLE);
  b.appendParagraph((mode==='final'?'FINAL':'DRAFT')+' • Last updated '+formatDateTime_(now)).setHeading(DocumentApp.ParagraphHeading.SUBTITLE);
  addInfoTable_(b,[['Project',d.projectTitle],['Project No.',d.projectNo],['Student / Team',d.studentTeam],['Email',d.email],['Date Submitted',d.dateSubmitted]]);
  addSection_(b,'1. Business Case',d.businessCase); addSection_(b,'2. Project Description',d.projectDescription); addSection_(b,'3. Project Goals',d.goals); addSection_(b,'4. Project Objectives',d.objectives);
  addArraySection_(b,'5. Project Deliverables',d.deliverables,['Deliverable No.','Description']); addArraySection_(b,'6. Project Tasks',d.tasks,['Task No.','Description','For Deliverable No.']);
  addSection_(b,'7. Out of Scope',d.outOfScope); addSection_(b,'8. WBS Link',d.wbsLink); addArraySection_(b,'9. Work Breakdown Structure',d.wbs,['Task No.','Description','For Deliverable No.']);
  addSection_(b,'10. Cost, Scheduling, and Quality Targets',d.targets); addArraySection_(b,'11. Project Risks',d.risks,['Issue/Risk','Likelihood','Impact','Description / Trigger','Risk Response Strategy']); addArraySection_(b,'12. Project Assumptions',d.assumptions,['No.','Assumption','Impact if False']);
  b.appendParagraph('13. Project Constraints').setHeading(DocumentApp.ParagraphHeading.HEADING1); addInfoTable_(b,[['Project Start',d.projectStart],['Go-Live',d.goLive],['Project End',d.projectEnd],['Estimated Hours',d.estimatedHours],['Hard Deadlines',d.deadlines],['Budget Constraints',d.budgetConstraints],['Quality / Performance Constraints',d.qualityConstraints],['Equipment / Personnel Constraints',d.equipmentConstraints],['Regulatory Constraints',d.regulatoryConstraints]]);
  addArraySection_(b,'14. Project Milestone Schedule',d.milestones,['Deliverable No.','Description','Date']); addArraySection_(b,'15. Contact List',d.contacts,['Name','Role','Email','Phone']); addSection_(b,'16. Comments',d.comments);
  b.appendParagraph('17. Approval').setHeading(DocumentApp.ParagraphHeading.HEADING1); addInfoTable_(b,[['Student / Team Approval',d.studentApproval],['Approval Date',d.approvalDate]]); doc.saveAndClose(); return doc.getId();
}
function upsertScheduleSheet_(fileId,d,mode,now){
  const folder=teamFolder_(d),name=clean_(d.projectTitle)+' - '+clean_(d.studentTeam)+' - Project Schedule'; let ss;
  try{ss=fileId?SpreadsheetApp.openById(fileId):null;}catch(e){ss=null;}
  if(!ss){ss=SpreadsheetApp.create(name);moveToFolder_(DriveApp.getFileById(ss.getId()),folder);} else ss.rename(name);
  const sh=ss.getSheets()[0];sh.clear();sh.setName('Project Schedule');
  sh.getRange('A1:H1').merge().setValue('PROJECT SCHEDULE').setFontWeight('bold').setFontSize(16).setHorizontalAlignment('center');
  sh.getRange('A2:B6').setValues([['Project',d.projectTitle||''],['Student / Team',d.studentTeam||''],['Email',d.email||''],['Status',mode==='final'?'FINAL':'DRAFT'],['Last Updated',formatDateTime_(now)]]);
  const headers=['Task No','Description','Dependencies','Plan Start','Plan Finish','Actual Start','Actual Finish','No. Days']; sh.getRange(8,1,1,headers.length).setValues([headers]).setFontWeight('bold');
  const rows=(d.schedule||[]).map(r=>headers.map(h=>r[h]||'')); if(rows.length)sh.getRange(9,1,rows.length,headers.length).setValues(rows);
  sh.setFrozenRows(8); sh.setColumnWidth(1,90);sh.setColumnWidth(2,340);sh.setColumnWidth(3,120);for(let c=4;c<=7;c++)sh.setColumnWidth(c,125);sh.setColumnWidth(8,90);
  sh.getRange(8,1,Math.max(rows.length+1,2),headers.length).setVerticalAlignment('top').setWrap(true); return ss.getId();
}
function moveToFolder_(file,folder){folder.addFile(file);try{DriveApp.getRootFolder().removeFile(file);}catch(e){}}
function addSection_(b,title,text){b.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING1);b.appendParagraph(String(text||''));}
function addInfoTable_(b,pairs){const rows=pairs.filter(x=>x[1]).map(x=>[String(x[0]),String(x[1])]);if(rows.length)b.appendTable(rows);}
function addArraySection_(b,title,arr,headers){b.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING1);arr=Array.isArray(arr)?arr:[];if(!arr.length){b.appendParagraph('No entries.');return;}const rows=[headers].concat(arr.map(o=>headers.map(h=>String(o[h]||''))));const t=b.appendTable(rows);for(let i=0;i<headers.length;i++)t.getRow(0).getCell(i).editAsText().setBold(true);}
function formatDateTime_(iso){return Utilities.formatDate(new Date(iso),Session.getScriptTimeZone()||'America/New_York','MMM d, yyyy h:mm a');}
function json_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}
