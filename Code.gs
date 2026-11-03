/**
 * CALENDRA | Google Apps Script metadata service.
 * Web App: Execute as ME, Who has access: ANYONE.
 * Script Properties: SPREADSHEET_ID, FIREBASE_API_KEY, ALLOWED_ORIGINS
 * Example ALLOWED_ORIGINS=https://your-app.vercel.app,http://localhost:5500
 * Run calendraSetup() once in the editor, grant Spreadsheet/UrlFetch permissions.
 * IMPORTANT: Firebase Identity Toolkit accounts:lookup validates the Firebase ID token.
 * No Google Calendar user token is sent to this service or stored in Sheets.
 */
const CALENDRA_DB = Object.freeze({
  Users:['uid','email','name','photoURL','bio','createdAt','updatedAt'],
  UserSettings:['uid','json','updatedAt'],
  CalendarPreferences:['id','uid','calendarId','visible','color','updatedAt'],
  EventTemplates:['id','uid','title','duration','description','reminder','createdAt','updatedAt'],
  CustomCategories:['id','uid','name','color','createdAt','updatedAt'],
  NotificationHistory:['id','uid','title','detail','read','createdAt'],
  ActivityLogs:['id','uid','action','detail','createdAt'],
  AppMetadata:['key','value','updatedAt']
});
const CALENDRA_CHANNEL = 'CALENDRA_METADATA_RPC_V1';
const CALENDRA_ALLOWED_ACTIONS = Object.freeze([
  'system.health','system.bootstrap','user.profile.get','user.profile.update','user.profile.delete',
  'user.settings.get','user.settings.update','calendar.preferences.get','calendar.preferences.save',
  'eventTemplates.list','eventTemplates.create','eventTemplates.update','eventTemplates.delete',
  'categories.list','categories.create','categories.update','categories.delete',
  'notifications.list','notifications.markRead','activity.list'
]);

function calendraSetup() {
  const props=PropertiesService.getScriptProperties();
  const id=props.getProperty('SPREADSHEET_ID');
  if(!id)throw Error('Set Script Property SPREADSHEET_ID terlebih dahulu.');
  const ss=SpreadsheetApp.openById(id);
  Object.keys(CALENDRA_DB).forEach(table=>{
    let s=ss.getSheetByName(table);
    if(!s)s=ss.insertSheet(table);
    const columns=CALENDRA_DB[table];
    if(s.getLastRow()===0){s.getRange(1,1,1,columns.length).setValues([columns]);s.setFrozenRows(1);s.getRange(1,1,1,columns.length).setFontWeight('bold').setBackground('#ece5ff');}
    else if(s.getLastColumn()<columns.length){throw Error(table+' memiliki header yang tidak sesuai. Periksa sebelum migrasi.');}
  });
  return 'Setup complete: '+ss.getUrl();
}
function doGet(e){
  if(e&&e.parameter&&e.parameter.bridge==='1'){
    const nonce=String(e.parameter.nonce||'');
    const origins=calendraAllowedOrigins_();
    // JSON serialization protects values from becoming executable HTML.
    const payload=JSON.stringify({nonce:nonce.slice(0,128),origins:origins,channel:CALENDRA_CHANNEL}).replace(/</g,'\\u003c');
    const html=`<!doctype html><html><head><meta charset="utf-8"><title>Calendra Bridge</title></head>
<body><script>
(function(){'use strict';
 const c=${payload};const t=window.top;
 function send(v){t.postMessage(Object.assign({channel:c.channel,nonce:c.nonce},v),'*');}
 function allowed(origin){return c.origins.indexOf(origin)!==-1;}
 window.addEventListener('message',function(e){
   if(!allowed(e.origin))return;
   const m=e.data;if(!m||m.channel!==c.channel||m.nonce!==c.nonce)return;
   if(m.type==='ping'){send({type:'ready'});return;}
   if(m.type!=='request'||typeof m.id!=='string')return;
   google.script.run.withSuccessHandler(function(result){send({type:'response',id:m.id,result:result});})
     .withFailureHandler(function(err){send({type:'response',id:m.id,result:{success:false,error:{code:'RPC_FAILED',message:String(err&&err.message||'Server error')}}});})
     .calendraRpc({action:m.action,payload:m.payload||{},idToken:m.idToken||'',requestId:m.id});
 });
 send({type:'ready'});
})();
</script></body></html>`;
    return HtmlService.createHtmlOutput(html).setTitle('Calendra metadata bridge').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return ContentService.createTextOutput('Calendra metadata service is online. API calls use the authenticated HTML bridge.');
}
function doPost(){
  // Deliberately not exposing a public unauthenticated JSON or no-cors mutation endpoint.
  return ContentService.createTextOutput('Use authenticated HTML bridge.');
}
function calendraAllowedOrigins_(){
  const raw=PropertiesService.getScriptProperties().getProperty('ALLOWED_ORIGINS')||'';
  return raw.split(',').map(x=>x.trim()).filter(x=>/^https:\/\/[a-z0-9.-]+(?::\d{2,5})?$/.test(x)||/^http:\/\/localhost(?::\d{2,5})?$/.test(x));
}
function calendraReply_(data){return {success:true,data:data,error:null,timestamp:new Date().toISOString()};}
function calendraFailure_(code,message){return {success:false,data:null,error:{code,message},timestamp:new Date().toISOString()};}
function calendraProperties_(){return PropertiesService.getScriptProperties();}
function calendraSpreadsheet_(){const id=calendraProperties_().getProperty('SPREADSHEET_ID');if(!id)throw Error('SPREADSHEET_ID is missing');return SpreadsheetApp.openById(id);}
function calendraSheet_(table){if(!CALENDRA_DB[table])throw Error('Unknown table');const s=calendraSpreadsheet_().getSheetByName(table);if(!s)throw Error('Sheet missing: '+table+'; run calendraSetup');return s;}
function calendraRecords_(table){
  const sh=calendraSheet_(table);const rows=sh.getDataRange().getValues();if(rows.length<=1)return [];
  const headers=CALENDRA_DB[table];return rows.slice(1).map((values,i)=>({row:i+2,...Object.fromEntries(headers.map((c,k)=>[c,values[k]??'']))}));
}
function calendraPublic_(obj){if(!obj)return null;const copy=Object.assign({},obj);delete copy.row;return copy;}
function calendraWrite_(table,values,row){
  const sh=calendraSheet_(table),headers=CALENDRA_DB[table];
  const data=headers.map(col=>values[col]===undefined?'':values[col]);
  if(row)sh.getRange(row,1,1,data.length).setValues([data]);else sh.appendRow(data);
}
function calendraDelete_(table,row){calendraSheet_(table).deleteRow(row);}
function calendraShort_(x,max=500){return String(x??'').trim().slice(0,max);}
function calendraJSON_(obj){return JSON.stringify(obj||{}).slice(0,20000);}
function calendraUid_(){return Utilities.getUuid();}
function calendraNow_(){return new Date().toISOString();}
function calendraUser_(idToken){
  if(typeof idToken!=='string'||idToken.length<50||idToken.length>9000)throw Error('Missing or invalid Firebase ID token');
  const key=calendraProperties_().getProperty('FIREBASE_API_KEY');if(!key)throw Error('FIREBASE_API_KEY missing in Script Properties');
  const response=UrlFetchApp.fetch('https://identitytoolkit.googleapis.com/v1/accounts:lookup?key='+encodeURIComponent(key),{
    method:'post',contentType:'application/json',payload:JSON.stringify({idToken}),muteHttpExceptions:true
  });
  if(response.getResponseCode()!==200)throw Error('Unauthenticated: Firebase token invalid or expired');
  const json=JSON.parse(response.getContentText()),u=json.users&&json.users[0];
  if(!u||!u.localId||u.disabled||!u.emailVerified)throw Error('Account not verified or disabled');
  const uid=u.localId;if(!/^[\w-]{5,256}$/.test(uid))throw Error('Invalid UID');
  return {uid,email:u.email||'',name:u.displayName||'',photoURL:u.photoUrl||''};
}
function calendraLog_(uid,action,detail){calendraWrite_('ActivityLogs',{id:calendraUid_(),uid,action,detail:calendraShort_(detail,250),createdAt:calendraNow_()});}
function calendraNotify_(uid,title,detail){calendraWrite_('NotificationHistory',{id:calendraUid_(),uid,title,detail,read:false,createdAt:calendraNow_()});}
function calendraIsWrite_(action){return action==='user.profile.get'||/\.(create|update|delete|save|markRead)$/.test(action)||action==='system.bootstrap';}

// Accessible to google.script.run; every call verifies the actual Firebase ID token.
function calendraRpc(request){
  try{
    if(!request||typeof request!=='object'||!CALENDRA_ALLOWED_ACTIONS.includes(request.action))return calendraFailure_('UNKNOWN_ACTION','Unsupported action');
    const u=calendraUser_(request.idToken),a=request.action,p=request.payload||{};
    if(p===null||typeof p!=='object'||Array.isArray(p))return calendraFailure_('BAD_PAYLOAD','Payload must be an object');
    if(JSON.stringify(p).length>25000)return calendraFailure_('TOO_LARGE','Payload too large');
    let lock=null;
    if(calendraIsWrite_(a)){
      lock=LockService.getScriptLock();lock.waitLock(8000);
      const req=calendraShort_(request.requestId||'',100);
      if(req&&req.length>5){const cache=CacheService.getScriptCache(),key='c_req_'+u.uid+'_'+req;if(cache.get(key))return calendraFailure_('DUPLICATE','Duplicate request');cache.put(key,'1',600);}
    }
    try{return calendraReply_(calendraDispatch_(u,a,p));}
    finally{if(lock)lock.releaseLock();}
  }catch(err){
    const message=String(err&&err.message||'Unknown error');
    const safe=/^(Unauthenticated|Account not verified|Invalid UID|Missing or invalid|Missing|Unsupported|User profile|Not found|No permission|Name|Title|Duration|Invalid|Too many|FIREBASE_API_KEY|SPREADSHEET_ID|Sheet missing|Unknown|The coordinates)/i.test(message)?message:'Server error. Check Apps Script execution logs.';
    return calendraFailure_('RPC_ERROR',safe);
  }
}
function calendraDispatch_(user,a,p){
  const uid=user.uid,now=calendraNow_();
  const own=t=>calendraRecords_(t).filter(x=>String(x.uid)===uid);
  const find=(t,id)=>own(t).find(x=>x.id===id);
  if(a==='system.health')return {service:'Calendra metadata service',authorized:true,uid,timestamp:now};
  if(a==='system.bootstrap'||a==='user.profile.get'){
    let profile=calendraRecords_('Users').find(x=>x.uid===uid);
    if(!profile){profile={uid,email:user.email,name:user.name||user.email.split('@')[0],photoURL:user.photoURL,bio:'',createdAt:now,updatedAt:now};calendraWrite_('Users',profile);}
    return {profile:calendraPublic_(profile),settings:calendraSettings_(uid),templates:own('EventTemplates').map(calendraPublic_),categories:own('CustomCategories').map(calendraPublic_),preferences:own('CalendarPreferences').map(calendraPublic_)};
  }
  if(a==='user.profile.update'){
    let rec=calendraRecords_('Users').find(x=>x.uid===uid);
    if(!rec){rec={uid,email:user.email,createdAt:now};}
    const next={uid,email:user.email,name:calendraShort_(p.name||user.name||rec.name,100),photoURL:calendraShort_(p.photoURL??rec.photoURL,900),bio:calendraShort_(p.bio??rec.bio,1000),createdAt:rec.createdAt,updatedAt:now};
    calendraWrite_('Users',next,rec.row);calendraLog_(uid,a,'Profile updated');calendraNotify_(uid,'Profile updated','Your Calendra profile was changed.');return next;
  }
  if(a==='user.settings.get')return calendraSettings_(uid);
  if(a==='user.settings.update'){
    const settings=calendraSettings_(uid);
    const allowed=['timezone','language','dateFormat','timeFormat','firstDay','workStart','workEnd','defaultDuration','defaultCalendar','defaultReminder','goals','notifyUpcoming'];
    allowed.forEach(k=>{if(Object.prototype.hasOwnProperty.call(p,k))settings[k]=p[k];});
    if(!/^\d{2}:\d{2}$/.test(settings.workStart)||!/^\d{2}:\d{2}$/.test(settings.workEnd))throw Error('Invalid working hours');
    if(!Number.isFinite(Number(settings.defaultDuration))||Number(settings.defaultDuration)<5||Number(settings.defaultDuration)>1440)throw Error('Invalid duration');
    const rec=own('UserSettings')[0];calendraWrite_('UserSettings',{uid,json:calendraJSON_(settings),updatedAt:now},rec&&rec.row);calendraLog_(uid,a,'Settings updated');calendraNotify_(uid,'Preferences saved','Your workspace settings were updated.');return settings;
  }
  if(a==='calendar.preferences.get')return own('CalendarPreferences').map(calendraPublic_);
  if(a==='calendar.preferences.save'){
    const calendarId=calendraShort_(p.calendarId,250);if(!calendarId)throw Error('Missing calendar ID');
    const rec=own('CalendarPreferences').find(x=>x.calendarId===calendarId);
    const value={id:rec?.id||calendraUid_(),uid,calendarId,visible:p.visible===false?false:true,color:/^#[\da-fA-F]{6}$/.test(p.color||'')?p.color:'',updatedAt:now};
    calendraWrite_('CalendarPreferences',value,rec?.row);return value;
  }
  if(a==='eventTemplates.list')return own('EventTemplates').map(calendraPublic_);
  if(a==='eventTemplates.create'||a==='eventTemplates.update'){
    const old=a.endsWith('update')?find('EventTemplates',p.id):null;
    if(a.endsWith('update')&&!old)throw Error('Not found or no permission');
    const title=calendraShort_(p.title,140),duration=Number(p.duration||60);
    if(!title)throw Error('Title is required');if(duration<5||duration>1440||!Number.isInteger(duration))throw Error('Duration must be 5–1440 minutes');
    const r={id:old?.id||calendraUid_(),uid,title,duration,description:calendraShort_(p.description,2000),reminder:Math.min(40320,Math.max(0,Number(p.reminder)||0)),createdAt:old?.createdAt||now,updatedAt:now};
    calendraWrite_('EventTemplates',r,old?.row);calendraLog_(uid,a,title);calendraNotify_(uid,'Template saved',title);return r;
  }
  if(a==='eventTemplates.delete'){const old=find('EventTemplates',p.id);if(!old)throw Error('Not found or no permission');calendraDelete_('EventTemplates',old.row);calendraLog_(uid,a,old.title);return {deleted:true};}
  if(a==='categories.list')return own('CustomCategories').map(calendraPublic_);
  if(a==='categories.create'||a==='categories.update'){
    const old=a.endsWith('update')?find('CustomCategories',p.id):null;if(a.endsWith('update')&&!old)throw Error('Not found or no permission');
    const name=calendraShort_(p.name,80);if(!name)throw Error('Name required');const color=/^#[0-9a-fA-F]{6}$/.test(p.color||'')?p.color:'#8453f1';
    const r={id:old?.id||calendraUid_(),uid,name,color,createdAt:old?.createdAt||now,updatedAt:now};calendraWrite_('CustomCategories',r,old?.row);return r;
  }
  if(a==='categories.delete'){const old=find('CustomCategories',p.id);if(!old)throw Error('Not found or no permission');calendraDelete_('CustomCategories',old.row);return {deleted:true};}
  if(a==='notifications.list')return own('NotificationHistory').slice(-100).reverse().map(calendraPublic_);
  if(a==='notifications.markRead'){const r=find('NotificationHistory',p.id);if(!r)throw Error('Not found or no permission');r.read=true;calendraWrite_('NotificationHistory',r,r.row);return {read:true};}
  if(a==='activity.list')return own('ActivityLogs').slice(-100).reverse().map(calendraPublic_);
  if(a==='user.profile.delete'){
    const tables=['Users','UserSettings','CalendarPreferences','EventTemplates','CustomCategories','NotificationHistory','ActivityLogs'];
    tables.forEach(t=>{const rows=calendraRecords_(t).filter(x=>x.uid===uid).map(x=>x.row).sort((a,b)=>b-a);rows.forEach(r=>calendraDelete_(t,r));});
    return {deleted:true,note:'Only Calendra Sheets metadata was deleted; Google Calendar events remain unchanged.'};
  }
  throw Error('Unsupported action');
}
function calendraSettings_(uid){
  const r=calendraRecords_('UserSettings').find(x=>x.uid===uid);
  const defaults={timezone:'Asia/Jakarta',language:'id-ID',dateFormat:'long',timeFormat:'24',firstDay:1,workStart:'09:00',workEnd:'17:00',defaultDuration:60,defaultCalendar:'primary',defaultReminder:15,goals:'',notifyUpcoming:true};
  if(!r)return defaults;
  try{return Object.assign(defaults,JSON.parse(r.json||'{}'));}catch(e){return defaults;}
}
