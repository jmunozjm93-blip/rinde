'use strict';
(()=>{
 const dialog=document.createElement('dialog');dialog.id='excelDialog';dialog.style.width='min(700px,calc(100% - 24px))';
 dialog.innerHTML=`<form id="excelForm"><div class="eyebrow">Formato de Comercial Depor</div><h2>Descargar rendición en Excel</h2><p id="excelSummary"></p><p>Los datos del encabezado vienen del formulario y cada monto irá a la categoría guardada con su boleta.</p><div class="fields" id="excelHeader"></div><p id="legacyCategoryNote" hidden>Solo estos registros antiguos necesitan clasificación. Las nuevas boletas ya la guardan desde el formulario.</p><div id="excelRows"></div><p id="excelError" role="alert" style="color:#a32943"></p><div class="dialog-actions"><button class="secondary" type="button" id="excelClose">Cerrar</button><button class="primary" type="submit" id="excelPrepare">Preparar Excel</button><a class="primary" id="excelSave" hidden style="text-decoration:none;text-align:center">↓ Descargar Excel .xlsx</a></div></form>`;
 document.body.append(dialog);
 const fields=[['type','Tipo'],['number','N.º de rendición'],['period','Periodo'],['costCenter','Centro de costo'],['brand','Marca'],['channel','Canal'],['branch','Sucursal'],['sap','Código SAP'],['check','N.º de cheque'],['bank','Banco'],['manager','Jefe directo']];
 for(const [key,label]of fields){const options=window.RINDE_CATALOGS[key];const wrap=document.createElement('div'),lab=document.createElement('label'),input=document.createElement(key==='type'||options?'select':'input');lab.htmlFor=input.id='excel_'+key;lab.textContent=label;if(key==='type')for(const v of ['Reembolso','Fondo por rendir']){const o=document.createElement('option');o.value=o.textContent=v;input.append(o);}else if(options){const blank=document.createElement('option');blank.value='';blank.textContent='Selecciona '+label.toLowerCase();input.append(blank);for(const [code,name] of options){const o=document.createElement('option');o.value=code+' · '+name;o.textContent=name+' · '+code;input.append(o);}}else input.maxLength=100;wrap.append(lab,input);$('excelHeader').append(wrap);}
 const headerPanel=document.createElement('details');headerPanel.open=true;headerPanel.className='panel';headerPanel.style.marginBottom='20px';headerPanel.innerHTML='<summary style="font-weight:700;font-size:16px;color:var(--ink);cursor:pointer">Datos de la rendición para Excel</summary><p>Las listas vienen del formato de oficina. Se recuerdan por persona en este dispositivo; puedes cambiarlas antes de descargar.</p>';
 headerPanel.append($('excelHeader'));$('receiptForm').before(headerPanel);
 let headerPerson='';const headerKey=()=> 'rinde-excel-header:'+headerPerson;
 function loadHeader(){const person=state.person||'';if(person===headerPerson)return;headerPerson=person;let saved={};try{saved=JSON.parse(localStorage.getItem(headerKey())||'{}');}catch(e){}for(const [k]of fields)$('excel_'+k).value=saved[k]||'';if(!$('excel_type').value)$('excel_type').value=state.budget?'Fondo por rendir':'Reembolso';headerPanel.hidden=!person||person==='__all__';}
 headerPanel.addEventListener('change',()=>{if(!headerPerson)return;try{localStorage.setItem(headerKey(),JSON.stringify(Object.fromEntries(fields.map(([k])=>[k,$('excel_'+k).value]))));}catch(e){toast('No se pudieron recordar los datos del encabezado en este dispositivo.');}});
 const originalRender=render;render=function(){originalRender();loadHeader();};loadHeader();
 let snapshot=null,exportEpoch=0,fileUrl=null,working=false;
 function clearFile(){if(fileUrl)URL.revokeObjectURL(fileUrl);fileUrl=null;$('excelSave').hidden=true;$('excelSave').removeAttribute('href');}
 function lock(value){dialog.querySelectorAll('input,select').forEach(el=>el.disabled=value);headerPanel.querySelectorAll('input,select').forEach(el=>el.disabled=value);}
 function reset(){exportEpoch++;snapshot=null;clearFile();lock(false);$('excelRows').replaceChildren();$('excelForm').reset();$('excelSummary').textContent='';$('excelError').textContent='';working=false;$('excelPrepare').disabled=false;}
 $('excelClose').onclick=()=>dialog.close();dialog.addEventListener('close',reset);
 // Session cleanup also clears export details and any generated file.
 const previousClearSession=clearSession;clearSession=function(message){if(dialog.open)dialog.close();reset();return previousClearSession(message);};
 $('excelForm').addEventListener('input',clearFile);$('excelForm').addEventListener('change',clearFile);
 $('exportCsv').onclick=()=>{
  if(!cloudReady||saving){toast('Espera a que termine la conexión o el guardado.');return;}
  if(state.person==='__all__'){toast('Selecciona una persona arriba para descargar su rendición.');$('person').focus();return;}
  if(!state.receipts.length)return;reset();snapshot=JSON.parse(JSON.stringify(state));
  const dates=snapshot.receipts.map(r=>r.date).sort();if(!$('excel_period').value)$('excel_period').value=dates[0].split('-').reverse().join('/')+' al '+dates.at(-1).split('-').reverse().join('/');
  $('excelSummary').textContent=snapshot.person+' · '+snapshot.receipts.length+' boletas · '+money(snapshot.receipts.reduce((s,r)=>s+r.amount,0));
  $('legacyCategoryNote').hidden=true;snapshot.receipts.forEach((r,i)=>{const wrap=document.createElement('div');wrap.className='field';const label=document.createElement('label'),select=document.createElement('select');label.htmlFor=select.id='excel_category_'+i;label.textContent=`${i+1}. ${r.merchant} · N.º ${r.number} · ${money(r.amount)}`;select.required=true;const blank=document.createElement('option');blank.value='';blank.textContent='Selecciona la columna de gasto';select.append(blank);RindeExcel.categories.forEach((c,n)=>{const option=document.createElement('option');option.value=String(n);option.textContent=c[0]+' '+c[1];select.append(option);});const suggested=RindeExcel.suggest(r);select.value=suggested===null?'':String(suggested);wrap.hidden=suggested!==null;if(suggested===null)$('legacyCategoryNote').hidden=false;wrap.append(label,select);$('excelRows').append(wrap);});
  dialog.showModal();
 };
 $('excelForm').onsubmit=async e=>{
  e.preventDefault();if(working||!snapshot)return;const epoch=exportEpoch,session=sessionEpoch,data=snapshot;
  const choices=data.receipts.map((_,i)=>{const v=$('excel_category_'+i).value;return v===''?null:Number(v);});
  if(choices.some(v=>v===null)){$('excelError').textContent='Selecciona una columna para cada gasto.';return;}
  const meta=Object.fromEntries(fields.map(([k])=>[k,$('excel_'+k).value.trim()]));for(const k of ['costCenter','brand','channel','branch'])meta[k]=meta[k].split(' · ')[0];working=true;lock(true);clearFile();$('excelError').textContent='Preparando tu archivo…';$('excelPrepare').disabled=true;
  try{
   await loadScript('vendor/exceljs-4.4.0.min.js',()=>!!window.ExcelJS);
   const wb=await RindeExcel.build(window.ExcelJS,data,meta,choices);const buffer=await wb.xlsx.writeBuffer();
   if(epoch!==exportEpoch||session!==sessionEpoch)return;
   fileUrl=URL.createObjectURL(new Blob([buffer],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
   const link=$('excelSave');link.href=fileUrl;link.download=`rendicion-${data.person.replace(/[^\p{L}\d]+/gu,'-')}-${today()}.xlsx`;link.hidden=false;link.focus();$('excelError').textContent='Archivo listo. Pulsa Descargar Excel para guardarlo.';
  }catch(err){if(epoch===exportEpoch)$('excelError').textContent=err.message||'No se pudo generar el Excel. Intenta nuevamente.';}
  finally{if(epoch===exportEpoch){working=false;lock(false);$('excelPrepare').disabled=false;}}
 };
})();
