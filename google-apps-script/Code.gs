/** Google Apps Script backend for the Project Management GitHub Pages forms.
 * SETUP: Create a new Apps Script project, paste this file, set DRIVE_FOLDER_ID, deploy as Web App.
 */
const DRIVE_FOLDER_ID = 'PASTE_GOOGLE_DRIVE_FOLDER_ID_HERE';
const INDEX_FILE = 'project-management-form-index.json';
function doPost(e){try{const req=JSON.parse(e.postData.contents);if(req.action!=='save')throw new Error('Unsupported action');const result=saveRecord_(req.formType,req.mode,req.data||{});return json_({ok:true,...result});}catch(err){return json_({ok:false,error:String(err.message||err)});}}
function doGet(e){try{if(e.parameter.action!=='load')throw new Error('Unsupported action');return json_({ok:true,data:loadRecord_(e.parameter.formType,e.parameter.recordId)});}catch(err){return json_({ok:false,error:String(err.message||err)});}}
function folder_(){if(DRIVE_FOLDER_ID.startsWith('PASTE_'))throw new Error('Set DRIVE_FOLDER_ID first.');return DriveApp.getFolderById(DRIVE_FOLDER_ID);}
function index_(){const f=folder_(),it=f.getFilesByName(INDEX_FILE);if(!it.hasNext())return {};return JSON.parse(it.next().getBlob().getDataAsString()||'{}');}
function writeIndex_(idx){const f=folder_(),it=f.getFilesByName(INDEX_FILE),s=JSON.stringify(idx,null,2);if(it.hasNext())it.next().setContent(s);else f.createFile(INDEX_FILE,s,MimeType.PLAIN_TEXT);}
function saveRecord_(formType,mode,data){const f=folder_(),idx=index_();let id=data.recordId||Utilities.getUuid();const now=new Date().toISOString();const rec={recordId:id,formType,mode,updatedAt:now,data:{...data,recordId:id}};let fileId=idx[id];if(fileId){try{DriveApp.getFileById(fileId).setContent(JSON.stringify(rec,null,2));}catch(e){fileId=null;}}if(!fileId){const safe=(data.studentTeam||data.email||'student').replace(/[^a-z0-9_-]+/gi,'-');const file=f.createFile(formType+'_'+safe+'_'+id+'.json',JSON.stringify(rec,null,2),MimeType.PLAIN_TEXT);fileId=file.getId();idx[id]=fileId;writeIndex_(idx);}return {recordId:id,fileId,updatedAt:now};}
function loadRecord_(formType,id){if(!id)throw new Error('Missing record ID');const idx=index_();if(!idx[id])throw new Error('Record not found');const rec=JSON.parse(DriveApp.getFileById(idx[id]).getBlob().getDataAsString());if(rec.formType!==formType)throw new Error('Form type mismatch');return rec.data;}
function json_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}
