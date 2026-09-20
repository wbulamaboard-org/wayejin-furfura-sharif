/*******************************************************
 WAYEJIN FURFURA SHARIF — Member System API
 Google Apps Script Web App backend
 *******************************************************/

// IMPORTANT:
// 1) Bind this script to the Google Sheet used for members.
// 2) Run setup() once from the Apps Script editor.
// 3) Set ADMIN_KEY below to a strong secret before deploying.
// 4) Deploy as Web app: Execute as you, access for anyone.
// The website calls this web-app URL for registration/login/verification.

const ADMIN_KEY = 'CHANGE_THIS_TO_A_STRONG_ADMIN_KEY';
const SHEET_NAME = 'Members';

const HEADERS = [
  'created_at','application_no','member_id','name','mobile','district','area',
  'education','experience','address','password_hash','status','approved_at','approved_by'
];

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) sh.appendRow(HEADERS);
  else {
    const first = sh.getRange(1,1,1,Math.max(sh.getLastColumn(),HEADERS.length)).getValues()[0];
    if (first.join('') === '') sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
  }
  sh.setFrozenRows(1);
  return 'SETUP OK';
}

function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    const action = String(p.action || 'verify').toLowerCase();
    if (action === 'verify') return json_(verify_(p.id || p.application_no || ''));
    if (action === 'health') return json_({ok:true,service:'WAYEJIN Member API',time:new Date().toISOString()});
    return json_({ok:false,message:'Unsupported GET action'});
  } catch (err) {
    return json_({ok:false,message:String(err.message || err)});
  }
}

function doPost(e) {
  try {
    const p = (e && e.parameter) || {};
    const action = String(p.action || '').toLowerCase();
    if (action === 'register') return json_(register_(p));
    if (action === 'login') return json_(login_(p));
    if (action === 'pending') return json_(pending_(p.admin_key));
    if (action === 'approve') return json_(approve_(p.admin_key,p.application_no));
    if (action === 'reject') return json_(reject_(p.admin_key,p.application_no));
    return json_({ok:false,message:'Unsupported action'});
  } catch (err) {
    return json_({ok:false,message:String(err.message || err)});
  }
}

function register_(p) {
  const name = clean_(p.name), mobile = clean_(p.mobile), district = clean_(p.district);
  const area = clean_(p.area), password = String(p.password || '');
  if (!name || !mobile || !district || !area || password.length < 6) {
    return {ok:false,message:'Required fields missing or password is too short.'};
  }
  const sh = sheet_();
  const rows = data_(sh);
  for (let i=0;i<rows.length;i++) {
    if (String(rows[i][4]).replace(/\D/g,'') === mobile.replace(/\D/g,'') &&
        String(rows[i][11]).toUpperCase() !== 'REJECTED') {
      return {ok:false,message:'এই mobile number দিয়ে একটি আবেদন ইতিমধ্যে আছে।'};
    }
  }
  const appNo = nextApplicationNo_(rows);
  sh.appendRow([
    new Date(),appNo,'',name,mobile,district,area,
    clean_(p.education),clean_(p.experience),clean_(p.address),
    hash_(password),'PENDING','',''
  ]);
  return {ok:true,application_no:appNo,status:'PENDING'};
}

function login_(p) {
  const key = clean_(p.id || p.application_no);
  const password = String(p.password || '');
  if (!key || !password) return {ok:false,message:'Member ID এবং password দিন।'};
  const rows = data_(sheet_());
  for (let i=0;i<rows.length;i++) {
    const r = rows[i];
    const app = String(r[1] || '').trim();
    const mid = String(r[2] || '').trim();
    if (key.toLowerCase() === app.toLowerCase() || key.toLowerCase() === mid.toLowerCase()) {
      if (String(r[11]).toUpperCase() !== 'APPROVED') return {ok:false,message:'আপনার আবেদন এখনও APPROVED নয়।'};
      if (hash_(password) !== String(r[10])) return {ok:false,message:'Password সঠিক নয়।'};
      return {ok:true,record:publicRecord_(r)};
    }
  }
  return {ok:false,message:'Member record পাওয়া যায়নি।'};
}

function verify_(key) {
  key = clean_(key);
  if (!key) return {found:false};
  const rows = data_(sheet_());
  for (let i=0;i<rows.length;i++) {
    const r = rows[i];
    if (String(r[1]).toLowerCase() === key.toLowerCase() ||
        String(r[2]).toLowerCase() === key.toLowerCase()) {
      if (String(r[11]).toUpperCase() !== 'APPROVED') return {found:false,status:String(r[11]||'PENDING')};
      return {found:true,record:publicRecord_(r)};
    }
  }
  return {found:false};
}

function pending_(key) {
  if (key !== ADMIN_KEY) return {ok:false,message:'Unauthorized'};
  const rows = data_(sheet_());
  return {ok:true,records:rows.filter(r=>String(r[11]).toUpperCase()==='PENDING').map(publicAdminRecord_)};
}

function approve_(key,appNo) {
  if (key !== ADMIN_KEY) return {ok:false,message:'Unauthorized'};
  if (!appNo) return {ok:false,message:'Application No required'};
  const sh=sheet_(), values=data_(sh);
  for(let i=0;i<values.length;i++){
    if(String(values[i][1])===String(appNo)){
      if(String(values[i][11]).toUpperCase()==='APPROVED') return {ok:true,message:'Already approved',member_id:values[i][2]};
      const memberId=nextMemberId_(values);
      sh.getRange(i+2,3).setValue(memberId);
      sh.getRange(i+2,12).setValue('APPROVED');
      sh.getRange(i+2,13).setValue(new Date());
      sh.getRange(i+2,14).setValue('ADMIN');
      return {ok:true,application_no:appNo,member_id:memberId};
    }
  }
  return {ok:false,message:'Application not found'};
}

function reject_(key,appNo) {
  if (key !== ADMIN_KEY) return {ok:false,message:'Unauthorized'};
  const sh=sheet_(), values=data_(sh);
  for(let i=0;i<values.length;i++){
    if(String(values[i][1])===String(appNo)){
      sh.getRange(i+2,12).setValue('REJECTED');
      return {ok:true,application_no:appNo,status:'REJECTED'};
    }
  }
  return {ok:false,message:'Application not found'};
}

function publicRecord_(r) {
  return {
    application_no:String(r[1]||''),
    wazeen_id:String(r[2]||''),
    name:String(r[3]||''),
    mobile:String(r[4]||''),
    district:String(r[5]||''),
    area:String(r[6]||''),
    education:String(r[7]||''),
    status:String(r[11]||'')
  };
}
function publicAdminRecord_(r) {
  const x=publicRecord_(r);
  x.created_at=r[0]?new Date(r[0]).toISOString():'';
  x.experience=String(r[8]||''); x.address=String(r[9]||'');
  return x;
}
function sheet_() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  if(!ss) throw new Error('Bind this script to the members Google Sheet.');
  let sh=ss.getSheetByName(SHEET_NAME);
  if(!sh){setup();sh=ss.getSheetByName(SHEET_NAME);}
  return sh;
}
function data_(sh) {
  const last=sh.getLastRow();
  return last<2?[]:sh.getRange(2,1,last-1,HEADERS.length).getValues();
}
function nextApplicationNo_(rows) {
  let max=0; rows.forEach(r=>{const m=String(r[1]||'').match(/WEFS\/APP\/\d{4}\/(\d+)/);if(m)max=Math.max(max,Number(m[1]))});
  return 'WEFS/APP/'+new Date().getFullYear()+'/'+String(max+1).padStart(8,'0');
}
function nextMemberId_(rows) {
  let max=0; rows.forEach(r=>{const m=String(r[2]||'').match(/WEFS\/WZ\/\d{4}\/(\d+)/);if(m)max=Math.max(max,Number(m[1]))});
  return 'WEFS/WZ/'+new Date().getFullYear()+'/'+String(max+1).padStart(6,'0');
}
function clean_(v){return String(v||'').trim().slice(0,500);}
function hash_(s){const bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(s),Utilities.Charset.UTF_8);return bytes.map(b=>(b<0?b+256:b).toString(16).padStart(2,'0')).join('');}
function json_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}
