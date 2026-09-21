/*******************************************************
 WAYEJIN FURFURA SHARIF — Member System API
 Google Apps Script Web App backend
 *******************************************************/
const ADMIN_KEY='CHANGE_THIS_TO_A_STRONG_ADMIN_KEY';
const SHEET_NAME='Members';
const LIVE_LOCATION_SHEET='LiveLocations';
const LOCATION_HEADERS=['member_id','name','mobile','lat','lng','accuracy_m','updated_at','sharing'];
const HEADERS=['created_at','application_no','member_id','name','mobile','district','area','education','experience','address','password_hash','status','approved_at','approved_by'];

function setup(){const ss=SpreadsheetApp.getActiveSpreadsheet();let sh=ss.getSheetByName(SHEET_NAME);if(!sh)sh=ss.insertSheet(SHEET_NAME);if(sh.getLastRow()===0)sh.appendRow(HEADERS);sh.setFrozenRows(1);return'SETUP OK';}
function doGet(e){try{const p=(e&&e.parameter)||{},a=String(p.action||'verify').toLowerCase();if(a==='verify')return json_(verify_(p.id||p.application_no||''));if(a==='health')return json_({ok:true,service:'WAYEJIN Member API',time:new Date().toISOString()});if(a==='list')return json_(list_(p.q||''));if(a==='location_list')return json_(locationList_(p.admin_key));if(a==='location_get')return json_(locationGet_(p.admin_key,p.member_id||p.id||''));return json_({ok:false,message:'Unsupported GET action'});}catch(x){return json_({ok:false,message:String(x.message||x)})}}
function doPost(e){try{const p=(e&&e.parameter)||{},a=String(p.action||'').toLowerCase();if(a==='register')return json_(register_(p));if(a==='login')return json_(login_(p));if(a==='pending')return json_(pending_(p.admin_key));if(a==='approve')return json_(approve_(p.admin_key,p.application_no));if(a==='reject')return json_(reject_(p.admin_key,p.application_no));if(a==='location_update')return json_(locationUpdate_(p));if(a==='location_stop')return json_(locationStop_(p));return json_({ok:false,message:'Unsupported action'});}catch(x){return json_({ok:false,message:String(x.message||x)})}}
function register_(p){const name=clean_(p.name),mobile=clean_(p.mobile),district=clean_(p.district),area=clean_(p.area),password=String(p.password||'');if(!name||!mobile||!district||!area||password.length<6)return{ok:false,message:'Required fields missing or password is too short.'};const sh=sheet_(),rows=data_(sh);for(let i=0;i<rows.length;i++)if(String(rows[i][4]).replace(/\D/g,'')===mobile.replace(/\D/g,'')&&String(rows[i][11]).toUpperCase()!=='REJECTED')return{ok:false,message:'এই mobile number দিয়ে একটি আবেদন ইতিমধ্যে আছে।'};const appNo=nextApplicationNo_(rows);sh.appendRow([new Date(),appNo,'',name,mobile,district,area,clean_(p.education),clean_(p.experience),clean_(p.address),hash_(password),'PENDING','','']);return{ok:true,application_no:appNo,status:'PENDING'}}
function login_(p){const key=clean_(p.id||p.application_no),password=String(p.password||'');if(!key||!password)return{ok:false,message:'Member ID এবং password দিন।'};for(const r of data_(sheet_())){if(key.toLowerCase()===String(r[1]||'').toLowerCase()||key.toLowerCase()===String(r[2]||'').toLowerCase()){if(String(r[11]).toUpperCase()!=='APPROVED')return{ok:false,message:'আপনার আবেদন এখনও APPROVED নয়।'};if(hash_(password)!==String(r[10]))return{ok:false,message:'Password সঠিক নয়।'};return{ok:true,record:publicRecord_(r)}}}return{ok:false,message:'Member record পাওয়া যায়নি।'}}
function verify_(key){key=clean_(key);if(!key)return{found:false};for(const r of data_(sheet_())){if(String(r[1]).toLowerCase()===key.toLowerCase()||String(r[2]).toLowerCase()===key.toLowerCase()){if(String(r[11]).toUpperCase()!=='APPROVED')return{found:false,status:String(r[11]||'PENDING')};return{found:true,record:publicRecord_(r)}}}return{found:false}}
function list_(q){q=String(q||'').toLowerCase().trim();const rows=data_(sheet_()).filter(r=>String(r[11]).toUpperCase()==='APPROVED');const records=rows.filter(r=>!q||[r[2],r[3],r[5],r[6]].join(' ').toLowerCase().includes(q)).map(publicRecord_);return{ok:true,records}}
function pending_(key){if(key!==ADMIN_KEY)return{ok:false,message:'Unauthorized'};return{ok:true,records:data_(sheet_()).filter(r=>String(r[11]).toUpperCase()==='PENDING').map(publicAdminRecord_)}}
function approve_(key,appNo){if(key!==ADMIN_KEY)return{ok:false,message:'Unauthorized'};const sh=sheet_(),values=data_(sh);for(let i=0;i<values.length;i++)if(String(values[i][1])===String(appNo)){if(String(values[i][11]).toUpperCase()==='APPROVED')return{ok:true,message:'Already approved',member_id:values[i][2]};const id=nextMemberId_(values);sh.getRange(i+2,3).setValue(id);sh.getRange(i+2,12).setValue('APPROVED');sh.getRange(i+2,13).setValue(new Date());sh.getRange(i+2,14).setValue('ADMIN');return{ok:true,application_no:appNo,member_id:id}}return{ok:false,message:'Application not found'}}
function reject_(key,appNo){if(key!==ADMIN_KEY)return{ok:false,message:'Unauthorized'};const sh=sheet_(),values=data_(sh);for(let i=0;i<values.length;i++)if(String(values[i][1])===String(appNo)){sh.getRange(i+2,12).setValue('REJECTED');return{ok:true,application_no:appNo,status:'REJECTED'}}return{ok:false,message:'Application not found'}}
function publicRecord_(r){return{application_no:String(r[1]||''),wazeen_id:String(r[2]||''),name:String(r[3]||''),mobile:String(r[4]||''),district:String(r[5]||''),area:String(r[6]||''),education:String(r[7]||''),status:String(r[11]||'')}}
function publicAdminRecord_(r){const x=publicRecord_(r);x.created_at=r[0]?new Date(r[0]).toISOString():'';x.experience=String(r[8]||'');x.address=String(r[9]||'');return x}
function sheet_(){const ss=SpreadsheetApp.getActiveSpreadsheet();if(!ss)throw Error('Bind this script to the members Google Sheet.');let sh=ss.getSheetByName(SHEET_NAME);if(!sh){setup();sh=ss.getSheetByName(SHEET_NAME)}return sh}
function data_(sh){const n=sh.getLastRow();return n<2?[]:sh.getRange(2,1,n-1,HEADERS.length).getValues()}
function nextApplicationNo_(rows){let max=0;rows.forEach(r=>{const m=String(r[1]||'').match(/WEFS\/APP\/\d{4}\/(\d+)/);if(m)max=Math.max(max,Number(m[1]))});return'WEFS/APP/'+new Date().getFullYear()+'/'+String(max+1).padStart(8,'0')}
function nextMemberId_(rows){let max=0;rows.forEach(r=>{const m=String(r[2]||'').match(/WEFS\/WZ\/\d{4}\/(\d+)/);if(m)max=Math.max(max,Number(m[1]))});return'WEFS/WZ/'+new Date().getFullYear()+'/'+String(max+1).padStart(6,'0')}
function clean_(v){return String(v||'').trim().slice(0,500)}
function hash_(s){const b=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(s),Utilities.Charset.UTF_8);return b.map(x=>(x<0?x+256:x).toString(16).padStart(2,'0')).join('')}
function json_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON)}

function setupLocation_(){const ss=SpreadsheetApp.getActiveSpreadsheet();let sh=ss.getSheetByName(LIVE_LOCATION_SHEET);if(!sh)sh=ss.insertSheet(LIVE_LOCATION_SHEET);if(sh.getLastRow()===0)sh.appendRow(LOCATION_HEADERS);sh.setFrozenRows(1);return sh}
function locationData_(sh){const n=sh.getLastRow();return n<2?[]:sh.getRange(2,1,n-1,LOCATION_HEADERS.length).getValues()}
function findApprovedMember_(memberId){const key=String(memberId||'').trim().toLowerCase();for(const r of data_(sheet_()))if(String(r[2]||'').toLowerCase()===key&&String(r[11]||'').toUpperCase()==='APPROVED')return r;return null}
function locationUpdate_(p){
 const memberId=clean_(p.member_id),mobile=String(p.mobile||'').replace(/\D/g,''),lat=Number(p.lat),lng=Number(p.lng),accuracy=Number(p.accuracy||0);
 if(!memberId||!mobile||!isFinite(lat)||!isFinite(lng)||lat<-90||lat>90||lng<-180||lng>180)return{ok:false,message:'Location data সঠিক নয়।'};
 const member=findApprovedMember_(memberId);if(!member)return{ok:false,message:'Approved member পাওয়া যায়নি।'};
 if(String(member[4]||'').replace(/\D/g,'')!==mobile)return{ok:false,message:'Member authentication failed.'};
 const sh=setupLocation_(),rows=locationData_(sh),now=new Date(),lock=LockService.getScriptLock();lock.waitLock(10000);
 try{for(let i=0;i<rows.length;i++)if(String(rows[i][0]).toLowerCase()===memberId.toLowerCase()){sh.getRange(i+2,1,1,8).setValues([[memberId,String(member[3]||''),mobile,lat,lng,isFinite(accuracy)?accuracy:0,now,true]]);return{ok:true,updated_at:now.toISOString(),member_id:memberId}}sh.appendRow([memberId,String(member[3]||''),mobile,lat,lng,isFinite(accuracy)?accuracy:0,now,true]);return{ok:true,updated_at:now.toISOString(),member_id:memberId}}finally{lock.releaseLock()}
}
function locationStop_(p){
 const memberId=clean_(p.member_id),mobile=String(p.mobile||'').replace(/\D/g,''),member=findApprovedMember_(memberId);if(!member)return{ok:false,message:'Approved member পাওয়া যায়নি।'};if(String(member[4]||'').replace(/\D/g,'')!==mobile)return{ok:false,message:'Member authentication failed.'};
 const sh=setupLocation_(),rows=locationData_(sh);for(let i=0;i<rows.length;i++)if(String(rows[i][0]).toLowerCase()===memberId.toLowerCase()){sh.getRange(i+2,8).setValue(false);return{ok:true,sharing:false}}return{ok:true,sharing:false}
}
function locationList_(key){
 if(key!==ADMIN_KEY)return{ok:false,message:'Unauthorized'};const now=Date.now(),rows=locationData_(setupLocation_());
 return{ok:true,records:rows.filter(r=>r[7]===true||String(r[7]).toUpperCase()==='TRUE').map(r=>({member_id:String(r[0]||''),name:String(r[1]||''),lat:Number(r[3]),lng:Number(r[4]),accuracy_m:Number(r[5]||0),updated_at:r[6]?new Date(r[6]).toISOString():'',sharing:true,age_seconds:r[6]?Math.max(0,Math.floor((now-new Date(r[6]).getTime())/1000)):null}))}
}
function locationGet_(key,memberId){
 if(key!==ADMIN_KEY)return{ok:false,message:'Unauthorized'};memberId=clean_(memberId);const rows=locationData_(setupLocation_());
 for(const r of rows)if(String(r[0]).toLowerCase()===memberId.toLowerCase())return{ok:true,record:{member_id:String(r[0]||''),name:String(r[1]||''),lat:Number(r[3]),lng:Number(r[4]),accuracy_m:Number(r[5]||0),updated_at:r[6]?new Date(r[6]).toISOString():'',sharing:r[7]===true||String(r[7]).toUpperCase()==='TRUE'}};
 return{ok:false,message:'এই সদস্যের কোনো location record নেই।'}
}
