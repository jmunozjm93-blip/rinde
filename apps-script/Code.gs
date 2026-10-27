const BOOK_ID = '1h05YvycSqoLKNaOvLH-kX1nJrv2_G8GDF3mT7VeMXR4';
const PEOPLE = ['Antonia Risopatron','Eduardo Paredes','Francesca Carnevali','Francisco Taito','German Correa','Jonathan Muñoz','Pablo Herceg','Sofia Torres','Tabata Rendic','Victoria Sabugo'];

function json_(value){return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);}
function doGet(){return json_({ok:true,service:'Rinde API',version:2});}
function hash_(s){return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,s).map(b=>('0'+((b+256)%256).toString(16)).slice(-2)).join('');}
function users_(){const raw=PropertiesService.getScriptProperties().getProperty('RINDE_USERS');if(!raw)throw Error('El acceso de usuarios aún no está configurado.');return JSON.parse(raw);}
function authError_(message){const e=Error(message||'Tu sesión terminó. Ingresa nuevamente.');e.code='UNAUTHORIZED';return e;}
function sessionKey_(token){return 'session:'+hash_(token);}
function login_(arg){
 if(!arg||typeof arg.userId!=='string'||typeof arg.password!=='string'||arg.password.length>200)throw authError_('Nombre o clave incorrectos.');
 return locked_(()=>{
  const cache=CacheService.getScriptCache(),user=users_().find(u=>u.id===arg.userId),attemptKey='login:'+hash_(arg.userId),attempt=JSON.parse(cache.get(attemptKey)||'{"count":0,"until":0}');
  if(attempt.until>Date.now())throw Error('Demasiados intentos. Espera un minuto y vuelve a intentar.');
  if(!user||hash_(user.salt+':'+arg.password)!==user.digest){const count=attempt.count+1;cache.put(attemptKey,JSON.stringify({count,until:count>=5?Date.now()+60000:0}),60);throw authError_('Nombre o clave incorrectos.');}
  cache.remove(attemptKey);const token=Utilities.getUuid()+Utilities.getUuid(),expires=Date.now()+21600000;
  cache.put(sessionKey_(token),JSON.stringify({id:user.id,expires}),21600);
  return {token,expires,user:{id:user.id,name:user.name,admin:user.admin}};
 });
}
function authenticate_(token){
 if(typeof token!=='string'||token.length<40||token.length>100)throw authError_();
 const raw=CacheService.getScriptCache().get(sessionKey_(token));if(!raw)throw authError_();const session=JSON.parse(raw);
 if(session.expires<Date.now())throw authError_();const user=users_().find(u=>u.id===session.id);if(!user)throw authError_();return user;
}
function target_(user,person,allowAll){
 if(person==='__all__'&&user.admin&&allowAll)return person;
 const name=person||user.name;if(!PEOPLE.includes(name)||(!user.admin&&name!==user.name))throw Error('No tienes acceso a esa rendición.');return name;
}
function doPost(e){try{
 if(!e||!e.postData||e.postData.contents.length>6500000)throw Error('Solicitud inválida.');const body=JSON.parse(e.postData.contents);
 if(body.action==='login')return json_({ok:true,data:login_(body.arg)});
 const user=authenticate_(body.token);let result;
 switch(body.action){
 case 'getState':result=state_(book_(),user,target_(user,body.arg&&body.arg.person,true));break;
 case 'setBudget':result=setBudget(body.arg,user);break;
 case 'saveReceipt':result=saveReceipt(body.arg,user);break;
 case 'removeReceipt':result=removeReceipt(body.arg,user);break;
 case 'getPhoto':result=getPhoto_(body.arg,user);break;
 case 'logout':CacheService.getScriptCache().remove(sessionKey_(body.token));result={loggedOut:true};break;
 default:throw Error('Acción desconocida.');
 }
 return json_({ok:true,data:result});
 }catch(err){return json_({ok:false,code:err.code||'ERROR',message:err.message||'No se pudo completar la operación.'});}}
function book_(){return SpreadsheetApp.openById(BOOK_ID);}
function receipts_(book){return book.getSheets().find(s=>s.getSheetId()===0);}
function rows_(sheet){const n=sheet.getLastRow()-1;if(n>10000)throw Error('La planilla supera 10.000 registros.');return n>0?sheet.getRange(2,1,n,10).getValues():[];}
function date_(value){return value instanceof Date?Utilities.formatDate(value,'America/Santiago','yyyy-MM-dd'):String(value);}
function budgets_(book){const sheet=book.getSheetByName('Fondos');if(!sheet)throw Error('Falta configurar los fondos personales.');const result={};if(sheet.getLastRow()>1)sheet.getRange(2,1,sheet.getLastRow()-1,2).getValues().forEach(r=>{if(PEOPLE.includes(String(r[0])))result[String(r[0])]=Number(r[1])||0;});return result;}
function state_(book,user,person){
 const all=person==='__all__',funds=budgets_(book),rows=rows_(receipts_(book));
 return {version:1,user:{id:user.id,name:user.name,admin:user.admin},person,budget:all?Object.values(funds).reduce((a,b)=>a+b,0):funds[person]||0,receipts:rows.filter(r=>r[0]&&r[9]==='Vigente'&&(all||String(r[2])===person)).map(r=>({id:String(r[0]),person:String(r[2]),date:date_(r[3]),number:String(r[4]),merchant:String(r[5]),category:String(r[6]),amount:Number(r[7]),photoUrl:String(r[8]||'')}))};
}
function locked_(callback){const lock=LockService.getScriptLock();lock.waitLock(30000);try{return callback();}finally{lock.releaseLock();}}
function setBudget(arg,user){
 if(!arg)throw Error('Falta el fondo.');const person=target_(user,arg.person,false),amount=arg.amount;
 if(!Number.isSafeInteger(amount)||amount<=0||amount>100000000000)throw Error('Monto de fondo inválido.');
 return locked_(()=>{const book=book_(),sheet=book.getSheetByName('Fondos'),rows=sheet.getRange(2,1,sheet.getLastRow()-1,2).getValues(),i=rows.findIndex(r=>r[0]===person);if(i<0)throw Error('No se encontró el fondo personal.');sheet.getRange(i+2,2).setValue(amount);SpreadsheetApp.flush();return state_(book,user,person);});
}
function ownRow_(book,user,id){text_(id,100);const rows=rows_(receipts_(book)),index=rows.findIndex(r=>String(r[0])===id&&r[9]==='Vigente');if(index<0)throw Error('No se encontró la boleta.');target_(user,String(rows[index][2]),false);return {row:rows[index],index};}
function removeReceipt(arg,user){if(!arg)throw Error('Falta la boleta.');return locked_(()=>{const book=book_(),found=ownRow_(book,user,arg.id),person=target_(user,arg.person,true);receipts_(book).getRange(found.index+2,10).setValue('Anulada');SpreadsheetApp.flush();return state_(book,user,person);});}
function getPhoto_(arg,user){const book=book_(),found=ownRow_(book,user,arg&&arg.id),url=String(found.row[8]);const match=url.match(/\/d\/([A-Za-z0-9_-]+)/)||url.match(/[?&]id=([A-Za-z0-9_-]+)/);if(!match)throw Error('Esta boleta no tiene foto.');const blob=DriveApp.getFileById(match[1]).getBlob();if(blob.getContentType()!=='image/jpeg')throw Error('Formato de foto no compatible.');return {mimeType:'image/jpeg',data:Utilities.base64Encode(blob.getBytes())};}
function text_(value,max){if(typeof value!=='string'||!value.trim()||value.length>max)throw Error('Los datos de la boleta son inválidos.');return value.trim();}
function safe_(value){return /^[\s]*[=+\-@]/.test(value)?"'"+value:value;}
function blob_(attachment){
 if(!attachment||attachment.mimeType!=='image/jpeg'||typeof attachment.data!=='string'||attachment.data.length>6000000||!/^[A-Za-z0-9+/=]+$/.test(attachment.data))throw Error('La foto es inválida o demasiado grande.');
 const bytes=Utilities.base64Decode(attachment.data);if(bytes.length<4||(bytes[0]&255)!==255||(bytes[1]&255)!==216)throw Error('La foto no es un JPEG válido.');
 return Utilities.newBlob(bytes,'image/jpeg','boleta.jpg');
}
function photoFolder_(){
 const props=PropertiesService.getUserProperties();const id=props.getProperty('RINDE_PHOTOS');
 if(id){try{return DriveApp.getFolderById(id);}catch(e){}}
 const folder=DriveApp.createFolder('Rinde · Boletas');props.setProperty('RINDE_PHOTOS',folder.getId());return folder;
}
function saveReceipt(payload,user){
 if(!payload||!payload.receipt)throw Error('Faltan los datos de la boleta.');
 const r=payload.receipt;r.id=text_(r.id,100);r.number=text_(r.number,40);r.merchant=text_(r.merchant,100);r.category=text_(r.category,40);
 r.person=target_(user,r.person,false);
 if(typeof r.date!=='string'||!/^\d{4}-\d{2}-\d{2}$/.test(r.date)||isNaN(Date.parse(r.date))||new Date(r.date).toISOString().slice(0,10)!==r.date)throw Error('Fecha inválida.');
 if(!Number.isSafeInteger(r.amount)||r.amount<=0||r.amount>100000000000)throw Error('Total inválido.');
 const blob=payload.attachment?blob_(payload.attachment):null;
 return locked_(()=>{
  const book=book_(),sheet=receipts_(book),rows=rows_(sheet);
  if(rows.some(row=>String(row[0])===r.id)){const previous=rows.find(row=>String(row[0])===r.id);if(String(previous[2])!==r.person)throw Error('Identificador no disponible.');return state_(book,user,r.person);} // A retry never duplicates a saved expense.
  if(rows.length>=10000)throw Error('Se alcanzó el máximo de 10.000 registros.');
  if(rows.some(row=>row[9]==='Vigente'&&String(row[4]).toLowerCase()===r.number.toLowerCase()&&String(row[5]).toLowerCase()===r.merchant.toLowerCase()))throw Error('Ya existe esta boleta para el mismo comercio.');
  const budget=budgets_(book)[r.person]||0;if(!budget)throw Error('Define primero el fondo inicial.');
  if(payload.budgetSeen!==budget)throw Error('El fondo cambió desde que abriste la app. Actualiza los datos y vuelve a revisar el saldo.');
  let photo=null;
  try{
   if(blob){photo=photoFolder_().createFile(blob).setName('Boleta '+r.number+' - '+r.person+' - '+r.date+'.jpg');}
   const row=sheet.getLastRow()+1;sheet.getRange(row,4,1,2).setNumberFormat('@');
   sheet.getRange(row,1,1,10).setValues([[safe_(r.id),new Date().toISOString(),r.person,r.date,safe_(r.number),safe_(r.merchant),safe_(r.category),r.amount,photo?photo.getUrl():'','Vigente']]);
   SpreadsheetApp.flush();
  }catch(e){if(photo){try{photo.setTrashed(true);}catch(ignore){}}throw e;}
  return state_(book,user,r.person);
 });
}
