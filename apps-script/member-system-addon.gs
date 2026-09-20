/*******************************************************
 WAYEJIN FURFURA SHARIF — MEMBER LOGIN / ID / QR ADD-ON
 This is intentionally additive: it uses the existing Applications
 sheet and existing status/wazeen_id fields. Do NOT replace the
 existing registration/admin code with this file.
 
 Integration:
 1) At the very TOP of the existing doPost(e), after it starts,
    add:
      if (WEFS_isMemberAction_(e)) return WEFS_memberPost_(e);
 2) At the very TOP of the existing doGet(e), add:
      if (WEFS_isMemberGet_(e)) return WEFS_memberGet_(e);
 3) Save, then deploy a NEW web-app version.
 
 Login credential: Wazeen ID + registered mobile number.
 Only APPROVED members can log in.
 *******************************************************/

function WEFS_isMemberAction_(e) {
  var p = e && e.parameter ? e.parameter : {};
  var a = String(p.action || '').toLowerCase();
  return a === 'member_login';
}

function WEFS_isMemberGet_(e) {
  var p = e && e.parameter ? e.parameter : {};
  var a = String(p.action || '').toLowerCase();
  return a === 'member_verify' || a === 'member_list' || a === 'member_health';
}

function WEFS_memberPost_(e) {
  try {
    var p = e && e.parameter ? e.parameter : {};
    var id = WEFS_clean_(p.id || p.wazeen_id || '');
    var mobile = WEFS_digits_(p.mobile || '');
    if (!id || !mobile) return WEFS_json_({ok:false,message:'Member ID এবং মোবাইল নম্বর দিন।'});

    var sh = getSheet_();
    var values = sh.getDataRange().getValues();
    if (values.length < 2) return WEFS_json_({ok:false,message:'Member database-এ কোনো record নেই।'});

    var h = values[0];
    var idCol = WEFS_col_(h, 'wazeen_id');
    var mobileCol = WEFS_col_(h, 'mobile');
    var statusCol = WEFS_col_(h, 'status');
    var nameCol = WEFS_col_(h, 'name');
    var districtCol = WEFS_col_(h, 'district');
    var blockCol = WEFS_col_(h, 'block');
    var areaCol = WEFS_col_(h, 'area');
    var appCol = WEFS_col_(h, 'application_no');

    if (idCol < 0 || mobileCol < 0 || statusCol < 0) {
      return WEFS_json_({ok:false,message:'wazeen_id/mobile/status column পাওয়া যায়নি।'});
    }

    for (var r=1; r<values.length; r++) {
      var row=values[r];
      var rowId=WEFS_clean_(row[idCol] || '');
      var rowMobile=WEFS_digits_(row[mobileCol] || '');
      var status=String(row[statusCol] || '').trim().toUpperCase();

      if (rowId.toLowerCase() === id.toLowerCase()) {
        if (status !== 'APPROVED') {
          return WEFS_json_({ok:false,message:'এই সদস্যের আবেদন এখনও APPROVED নয়।',status:status});
        }
        if (rowMobile !== mobile) {
          return WEFS_json_({ok:false,message:'Member ID অথবা মোবাইল নম্বর সঠিক নয়।'});
        }

        return WEFS_json_({
          ok:true,
          record:{
            wazeen_id:rowId,
            application_no:appCol>=0 ? String(row[appCol] || '') : '',
            name:nameCol>=0 ? String(row[nameCol] || '') : '',
            mobile:rowMobile,
            district:districtCol>=0 ? String(row[districtCol] || '') : '',
            block:blockCol>=0 ? String(row[blockCol] || '') : '',
            area:areaCol>=0 ? String(row[areaCol] || '') : '',
            status:'APPROVED'
          }
        });
      }
    }
    return WEFS_json_({ok:false,message:'Member ID পাওয়া যায়নি।'});
  } catch(err) {
    return WEFS_json_({ok:false,message:String(err && err.message || err)});
  }
}

function WEFS_memberGet_(e) {
  try {
    var p=e && e.parameter ? e.parameter : {};
    var a=String(p.action || '').toLowerCase();
    if (a==='member_health') {
      return WEFS_json_({ok:true,service:'WAYEJIN Member System',time:new Date().toISOString()});
    }

    var sh=getSheet_();
    var values=sh.getDataRange().getValues();
    if (values.length<2) return WEFS_json_({found:false,records:[]});

    var h=values[0];
    var idCol=WEFS_col_(h,'wazeen_id');
    var appCol=WEFS_col_(h,'application_no');
    var mobileCol=WEFS_col_(h,'mobile');
    var statusCol=WEFS_col_(h,'status');
    var nameCol=WEFS_col_(h,'name');
    var districtCol=WEFS_col_(h,'district');
    var blockCol=WEFS_col_(h,'block');
    var areaCol=WEFS_col_(h,'area');

    if (a==='member_verify') {
      var key=WEFS_clean_(p.id || p.wazeen_id || p.application_no || '');
      if (!key) return WEFS_json_({found:false});
      for(var r=1;r<values.length;r++){
        var row=values[r], status=String(row[statusCol]||'').trim().toUpperCase();
        var wid=idCol>=0?String(row[idCol]||'').trim():'';
        var app=appCol>=0?String(row[appCol]||'').trim():'';
        if((wid && wid.toLowerCase()===key.toLowerCase())||(app && app.toLowerCase()===key.toLowerCase())){
          if(status!=='APPROVED') return WEFS_json_({found:false,status:status});
          return WEFS_json_({found:true,record:{
            wazeen_id:wid,application_no:app,
            name:nameCol>=0?String(row[nameCol]||''):'',
            mobile:mobileCol>=0?String(row[mobileCol]||''):'',
            district:districtCol>=0?String(row[districtCol]||''):'',
            block:blockCol>=0?String(row[blockCol]||''):'',
            area:areaCol>=0?String(row[areaCol]||''):'',
            status:'APPROVED'
          }});
        }
      }
      return WEFS_json_({found:false});
    }

    if (a==='member_list') {
      var q=WEFS_clean_(p.q||'').toLowerCase();
      var records=[];
      for(var i=1;i<values.length;i++){
        var x=values[i];
        if(String(x[statusCol]||'').trim().toUpperCase()!=='APPROVED') continue;
        var rec={
          wazeen_id:idCol>=0?String(x[idCol]||''):'',
          application_no:appCol>=0?String(x[appCol]||''):'',
          name:nameCol>=0?String(x[nameCol]||''):'',
          district:districtCol>=0?String(x[districtCol]||''):'',
          block:blockCol>=0?String(x[blockCol]||''):'',
          area:areaCol>=0?String(x[areaCol]||''):'',
          status:'APPROVED'
        };
        var hay=(rec.wazeen_id+' '+rec.name+' '+rec.district+' '+rec.block+' '+rec.area).toLowerCase();
        if(!q || hay.indexOf(q)>=0) records.push(rec);
      }
      return WEFS_json_({ok:true,records:records});
    }

    return WEFS_json_({ok:false,message:'Unsupported member action'});
  } catch(err2) {
    return WEFS_json_({ok:false,message:String(err2 && err2.message || err2)});
  }
}

function WEFS_col_(headers,name) {
  var target=String(name).toLowerCase();
  for(var i=0;i<headers.length;i++){
    if(String(headers[i]||'').trim().toLowerCase()===target) return i;
  }
  return -1;
}

function WEFS_digits_(v) {
  return String(v||'').replace(/\D/g,'').slice(-10);
}

function WEFS_clean_(v) {
  return String(v||'').trim().slice(0,500);
}

function WEFS_json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
