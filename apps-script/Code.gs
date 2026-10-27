const BOOK_ID = '1h05YvycSqoLKNaOvLH-kX1nJrv2_G8GDF3mT7VeMXR4';
const PEOPLE = ['Antonia Risopatron','Eduardo Paredes','Francesca Carnevali','Francisco Taito','German Correa','Jonathan Muñoz','Pablo Herceg','Sofia Torres','Tabata Rendic','Victoria Sabugo'];
function doGet(){return json_({ok:true,service:'Rinde API',message:'Abre la app de GitHub Pages para ingresar.'});}
function json_(value){return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);}
function doPost(e){
 try{
  if(!e||!e.postData||e.postData.contents.length>6500000)return json_({ok:false,message:'Solicitud inválida.'});
  const body=JSON.parse(e.postData.contents),expected=PropertiesService.getScriptProperties().getProperty('ACCESS_DIGEST');
  if(!expected||typeof body.key!=='string'||body.key.length>200)return json_({ok:false,code:'UNAUTHORIZED',message:'Clave de acceso incorrecta.'});
  const digest=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,body.key).map(b=>('0'+((b+256)%256).toString(16)).slice(-2)).join('');
  if(digest!==expected)return json_({ok:false,code:'UNAUTHORIZED',message:'Clave de acceso incorrecta.'});
  const actions={getState:getState,setBudget:setBudget,saveReceipt:saveReceipt,removeReceipt:removeReceipt};
  if(!Object.prototype.hasOwnProperty.call(actions,body.action))return json_({ok:false,message:'Acción desconocida.'});
  return json_({ok:true,data:actions[body.action](body.arg)});
 }catch(err){return json_({ok:false,message:err.message||'No se pudo completar la operación.'});}
}
function book_(){ return SpreadsheetApp.openById(BOOK_ID); }
function receipts_(book){return book.getSheets().find(s=>s.getSheetId()===0);}
function rows_(sheet){const count=sheet.getLastRow()-1;if(count>10000)throw Error('La planilla supera 10.000 registros.');return count>0?sheet.getRange(2,1,count,10).getValues():[];}
function date_(value){return value instanceof Date?Utilities.formatDate(value,'America/Santiago','yyyy-MM-dd'):String(value);}
function state_(book){
 const rows=rows_(receipts_(book));
 return {version:1,budget:Number(book.getSheetByName('Configuración').getRange('B2').getValue())||0,receipts:rows.filter(r=>r[0]&&r[9]!=='Anulada').map(r=>({id:String(r[0]),person:String(r[2]),date:date_(r[3]),number:String(r[4]),merchant:String(r[5]),category:String(r[6]),amount:Number(r[7]),photoUrl:String(r[8]||'')}))};
}
function getState(){return state_(book_());}
function locked_(callback){const lock=LockService.getScriptLock();lock.waitLock(30000);try{return callback();}finally{lock.releaseLock();}}
function setBudget(amount){
 if(!Number.isSafeInteger(amount)||amount<=0||amount>100000000000)throw Error('Monto de fondo inválido.');
 return locked_(()=>{const book=book_();book.getSheetByName('Configuración').getRange('B2').setValue(amount);SpreadsheetApp.flush();return state_(book);});
}
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
function saveReceipt(payload){
 if(!payload||!payload.receipt)throw Error('Faltan los datos de la boleta.');
 const r=payload.receipt;r.id=text_(r.id,100);r.number=text_(r.number,40);r.merchant=text_(r.merchant,100);r.category=text_(r.category,40);
 if(!PEOPLE.includes(r.person))throw Error('Selecciona una persona del listado.');
 if(typeof r.date!=='string'||!/^\d{4}-\d{2}-\d{2}$/.test(r.date)||isNaN(Date.parse(r.date))||new Date(r.date).toISOString().slice(0,10)!==r.date)throw Error('Fecha inválida.');
 if(!Number.isSafeInteger(r.amount)||r.amount<=0||r.amount>100000000000)throw Error('Total inválido.');
 const blob=payload.attachment?blob_(payload.attachment):null;
 return locked_(()=>{
  const book=book_(),sheet=receipts_(book),rows=rows_(sheet);
  if(rows.some(row=>String(row[0])===r.id))return state_(book); // A retry never duplicates a saved expense.
  if(rows.length>=10000)throw Error('Se alcanzó el máximo de 10.000 registros.');
  if(rows.some(row=>row[9]!=='Anulada'&&String(row[4]).toLowerCase()===r.number.toLowerCase()&&String(row[5]).toLowerCase()===r.merchant.toLowerCase()))throw Error('Ya existe esta boleta para el mismo comercio.');
  const budget=Number(book.getSheetByName('Configuración').getRange('B2').getValue());if(!budget)throw Error('Define primero el fondo inicial.');
  if(payload.budgetSeen!==budget)throw Error('El fondo cambió desde que abriste la app. Actualiza los datos y vuelve a revisar el saldo.');
  let photo=null;
  try{
   if(blob){photo=photoFolder_().createFile(blob).setName('Boleta '+r.number+' - '+r.person+' - '+r.date+'.jpg');}
   const row=sheet.getLastRow()+1;sheet.getRange(row,4,1,2).setNumberFormat('@');
   sheet.getRange(row,1,1,10).setValues([[safe_(r.id),new Date().toISOString(),r.person,r.date,safe_(r.number),safe_(r.merchant),safe_(r.category),r.amount,photo?photo.getUrl():'','Vigente']]);
   SpreadsheetApp.flush();
  }catch(e){if(photo){try{photo.setTrashed(true);}catch(ignore){}}throw e;}
  return state_(book);
 });
}
function removeReceipt(id){
 text_(id,100);return locked_(()=>{const book=book_(),sheet=receipts_(book),rows=rows_(sheet),index=rows.findIndex(r=>String(r[0])===id);if(index<0)throw Error('No se encontró la boleta. Actualiza los datos.');sheet.getRange(index+2,10).setValue('Anulada');SpreadsheetApp.flush();return state_(book);});
}
