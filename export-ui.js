'use strict';
(()=>{
 const dialog=document.createElement('dialog');dialog.id='excelDialog';dialog.style.width='min(700px,calc(100% - 24px))';
 dialog.innerHTML=`<form id="excelForm"><div class="eyebrow">Formato de Comercial Depor</div><h2>Descargar rendición</h2><p id="excelSummary"></p><p>Se descargan juntos el Excel y el PDF con las fotos de las boletas (tamaño carta, 3 por hoja). Los datos del encabezado vienen del formulario y cada monto irá a la categoría guardada con su boleta.</p><div class="fields" id="excelHeader"></div><p id="legacyCategoryNote" hidden>Solo estos registros antiguos necesitan clasificación. Las nuevas boletas ya la guardan desde el formulario.</p><div id="excelRows"></div><p id="excelError" role="alert" style="color:#a32943"></p><div class="dialog-actions"><button class="secondary" type="button" id="excelClose">Cerrar</button><button class="primary" type="submit" id="excelPrepare">Preparar descarga</button><a class="primary" id="excelSave" hidden style="text-decoration:none;text-align:center">↓ Descargar Excel .xlsx</a><a class="secondary" id="excelPdfSave" hidden style="text-decoration:none;text-align:center">↓ Descargar PDF de boletas</a></div></form>`;
 document.body.append(dialog);
 const fields=[['type','Tipo'],['number','N.º de rendición'],['period','Periodo'],['costCenter','Centro de costo'],['brand','Marca'],['channel','Canal'],['branch','Sucursal'],['sap','Código SAP'],['check','N.º de cheque'],['bank','Banco'],['manager','Jefe directo']];
 // Valores predeterminados del formato de oficina; los campos sin dato quedan ocultos (no se borran).
 const defaults={type:'Fondo por rendir',number:'6',period:'DIC-FEB',costCenter:'2302 · Ventas Mayoristas Sell Out',brand:'TODAS · Todas las marcas',channel:'MAY · Mayoristas',branch:'IND4 · INDIRECTO CENTRAL'};
 const hiddenFields=['sap','check','bank'];
 for(const [key,label]of fields){const options=window.RINDE_CATALOGS[key];const wrap=document.createElement('div'),lab=document.createElement('label'),input=document.createElement(key==='type'||options?'select':'input');lab.htmlFor=input.id='excel_'+key;lab.textContent=label;if(key==='type')for(const v of ['Reembolso','Fondo por rendir']){const o=document.createElement('option');o.value=o.textContent=v;input.append(o);}else if(options){const blank=document.createElement('option');blank.value='';blank.textContent='Selecciona '+label.toLowerCase();input.append(blank);for(const [code,name] of options){const o=document.createElement('option');o.value=code+' · '+name;o.textContent=name+' · '+code;input.append(o);}}else input.maxLength=100;wrap.hidden=hiddenFields.includes(key);wrap.append(lab,input);$('excelHeader').append(wrap);}
 const headerPanel=document.createElement('details');headerPanel.open=true;headerPanel.className='panel';headerPanel.style.marginBottom='20px';headerPanel.innerHTML='<summary style="font-weight:700;font-size:16px;color:var(--ink);cursor:pointer">Datos de la rendición para Excel</summary><p>Las listas vienen del formato de oficina. Se recuerdan por persona en este dispositivo; puedes cambiarlas antes de descargar.</p>';
 headerPanel.append($('excelHeader'));$('receiptForm').before(headerPanel);
 let headerPerson='';const headerKey=()=> 'rinde-excel-header:'+headerPerson;
 function loadHeader(){const person=state.person||'';if(person===headerPerson)return;headerPerson=person;let saved={};try{saved=JSON.parse(localStorage.getItem(headerKey())||'{}');}catch(e){}for(const [k]of fields)$('excel_'+k).value=saved[k]||defaults[k]||'';if(!$('excel_type').value)$('excel_type').value=state.budget?'Fondo por rendir':'Reembolso';headerPanel.hidden=!person||person==='__all__';}
 headerPanel.addEventListener('change',()=>{if(!headerPerson)return;try{localStorage.setItem(headerKey(),JSON.stringify(Object.fromEntries(fields.map(([k])=>[k,$('excel_'+k).value]))));}catch(e){toast('No se pudieron recordar los datos del encabezado en este dispositivo.');}});
 const history=document.createElement('details');history.className='panel';history.style.marginTop='20px';history.innerHTML='<summary style="font-weight:700;cursor:pointer">Rendiciones pagadas</summary><p>Volver a descargar un Excel del historial no registra otro pago.</p><div id="paidExports"></div>';$('boletas').after(history);
 function renderHistory(){const box=$('paidExports');box.replaceChildren();const items=state.paidExports||[];history.hidden=!items.length;for(const batch of items){const btn=document.createElement('button');btn.type='button';btn.className='secondary full';btn.style.marginTop='8px';btn.textContent=batch.person+' · '+new Date(batch.at).toLocaleDateString('es-CL')+' · '+batch.count+' boletas · '+money(batch.total)+' · Volver a descargar';btn.onclick=()=>openPaid(batch.id);box.append(btn);}}
 const originalRender=render;render=function(){originalRender();loadHeader();renderHistory();};loadHeader();renderHistory();
 let snapshot=null,exportEpoch=0,fileUrl=null,working=false,prepared=null,committed=false,paying=false,pdfUrl=null,pdfName='',pdfNote='';
 // PDF de fotos de la rendición: se arma junto al Excel (antes de pagar, mientras las boletas siguen vigentes) y se descarga con el mismo botón.
 async function preparePdf(receipts,meta,person,epoch){
  const title='Rendición'+(meta.number?' N.º '+meta.number:'')+(meta.period?' · '+meta.period:'')+' · '+person;
  let res;try{res=await RindePdf.build({receipts,title,progress:t=>{if(epoch===exportEpoch)$('excelError').textContent=t;}});}
  catch(e){if(e.message==='La sesión cambió.')throw e;res={blob:null,failed:receipts.length};}
  if(epoch!==exportEpoch)return;
  pdfNote=RindePdf.summary(res);
  if(res.blob){pdfUrl=URL.createObjectURL(res.blob);pdfName=`boletas-rendicion-${RindePdf.slug(meta.number||'')||'sin-numero'}-${RindePdf.slug(person)}-${today()}.pdf`;}
 }
 // iPhone/iPad descargan un solo archivo por toque: el segundo reemplaza al primero.
 const oneAtATime=/iP(hone|ad|od)/.test(navigator.userAgent)||(navigator.platform==='MacIntel'&&navigator.maxTouchPoints>1);
 function savePdf(){if(!pdfUrl)return;const a=document.createElement('a');a.href=pdfUrl;a.download=pdfName;document.body.append(a);a.click();a.remove();}
 function showPdfLink(){const a=$('excelPdfSave');if(!pdfUrl){a.hidden=true;return;}a.href=pdfUrl;a.download=pdfName;a.hidden=false;}
 function clearFile(){if(fileUrl)URL.revokeObjectURL(fileUrl);fileUrl=null;if(pdfUrl)URL.revokeObjectURL(pdfUrl);pdfUrl=null;pdfName='';pdfNote='';prepared=null;committed=false;$('excelSave').hidden=true;$('excelSave').removeAttribute('href');$('excelPdfSave').hidden=true;$('excelPdfSave').removeAttribute('href');}
 function lock(value){dialog.querySelectorAll('input,select').forEach(el=>el.disabled=value);headerPanel.querySelectorAll('input,select').forEach(el=>el.disabled=value);}
 function reset(){exportEpoch++;snapshot=null;clearFile();lock(false);$('excelRows').replaceChildren();$('excelForm').reset();$('excelSummary').textContent='';$('excelError').textContent='';working=false;paying=false;$('excelClose').disabled=false;$('excelPrepare').disabled=false;$('excelPrepare').hidden=false;}
 dialog.addEventListener('cancel',e=>{if(paying)e.preventDefault();});
 $('excelClose').onclick=()=>dialog.close();dialog.addEventListener('close',()=>{if(!dialog.open)reset();});
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
   const paymentId=newId();const wb=await RindeExcel.build(window.ExcelJS,{...data,status:'Pagado',paymentId},meta,choices);const buffer=await wb.xlsx.writeBuffer();
   if(epoch!==exportEpoch||session!==sessionEpoch)return;
   const excelUrl=URL.createObjectURL(new Blob([buffer],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
   await preparePdf(data.receipts,meta,data.person,epoch);
   if(epoch!==exportEpoch||session!==sessionEpoch){URL.revokeObjectURL(excelUrl);return;}
   fileUrl=excelUrl;
   prepared={id:paymentId,person:data.person,budgetSeen:data.budget,ids:data.receipts.map(r=>r.id),expected:data.receipts.map(r=>[r.id,r.date,r.number,r.merchant,r.category,r.amount]),meta,choices};
   const link=$('excelSave');link.href=fileUrl;link.download=`rendicion-pagada-${data.person.replace(/[^\p{L}\d]+/gu,'-')}-${today()}.xlsx`;link.textContent=pdfUrl&&!oneAtATime?'↓ Descargar Excel + boletas y marcar pagado':'↓ Descargar Excel y marcar pagado';link.hidden=false;link.focus();$('excelError').textContent=pdfNote+' Al pulsar Descargar, estas boletas pasarán a Pagado y tu fondo quedará en cero. Se conservará una copia en el historial.'+(oneAtATime&&pdfUrl?' En iPhone se baja primero el Excel y luego aparece el botón del PDF.':'');
  }catch(err){if(epoch===exportEpoch)$('excelError').textContent=err.message||'No se pudo generar el Excel. Intenta nuevamente.';}
  finally{if(epoch===exportEpoch){working=false;lock(false);$('excelPrepare').disabled=false;}}
 };
 $('excelSave').onclick=async e=>{
  // Ya pagado: el enlace descarga el Excel y, a continuación, el PDF de boletas.
  if(committed){if(!oneAtATime)setTimeout(savePdf,800);return;}e.preventDefault();if(paying||!prepared)return;paying=true;working=true;lock(true);setSaving(true);$('excelClose').disabled=true;$('excelPrepare').disabled=true;$('excelError').textContent='Registrando el pago antes de descargar…';const epoch=exportEpoch,session=sessionEpoch;
  try{const next=await rpc('payExport',prepared);if(epoch!==exportEpoch||session!==sessionEpoch)return;state=next;committed=true;render();$('excelPrepare').hidden=true;$('excelSave').textContent='↓ Volver a descargar Excel';showPdfLink();$('excelError').textContent='Pagado. Fondo en cero. '+(!pdfUrl?'':oneAtATime?'Se descargó el Excel. En iPhone se baja un archivo a la vez: ahora toca “Descargar PDF de boletas”. ':'Se descargan el Excel y el PDF de boletas; si el navegador pregunta, permite descargar varios archivos. ')+'Si algún archivo no se abrió, vuelve a tocar su botón. También está en Rendiciones pagadas.';$('excelSave').click();}
  catch(err){if(epoch===exportEpoch)$('excelError').textContent=err.message+' Puedes reintentar con este mismo botón. Si el pago ya quedó registrado, no se repetirá.';}
  finally{if(epoch===exportEpoch){paying=false;working=false;$('excelClose').disabled=false;}setSaving(false);}
 };
 async function openPaid(id){
  reset();dialog.showModal();$('excelPrepare').hidden=true;working=true;$('legacyCategoryNote').hidden=true;$('excelError').textContent='Recuperando Excel pagado…';const epoch=exportEpoch,session=sessionEpoch;
  try{const data=await rpc('getPaidExport',{id});await loadScript('vendor/exceljs-4.4.0.min.js',()=>!!window.ExcelJS);const wb=await RindeExcel.build(window.ExcelJS,{...data,status:'Pagado'},data.meta,data.choices);const buffer=await wb.xlsx.writeBuffer();if(epoch!==exportEpoch||session!==sessionEpoch)return;
   const excelUrl=URL.createObjectURL(new Blob([buffer],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
   await preparePdf(data.receipts,data.meta||{},data.person,epoch);if(epoch!==exportEpoch||session!==sessionEpoch){URL.revokeObjectURL(excelUrl);return;}
   fileUrl=excelUrl;committed=true;const link=$('excelSave');link.href=fileUrl;link.download='rendicion-pagada-'+id+'.xlsx';link.textContent=pdfUrl&&!oneAtATime?'↓ Descargar copia: Excel + boletas':'↓ Descargar copia del Excel';link.hidden=false;showPdfLink();$('excelSummary').textContent=data.person+' · Pagado';$('excelError').textContent=pdfNote+' Esta es una copia del pago registrado. No cambia tu fondo ni vuelve a pagar las boletas.';
  }catch(err){if(epoch===exportEpoch)$('excelError').textContent=err.message;}finally{if(epoch===exportEpoch)working=false;}
 }
})();
