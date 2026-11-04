'use strict';
/* Descarga de las fotos de boletas en PDF tamaño carta, tres boletas por hoja. */
(()=>{
 const btn=document.createElement('button');btn.type='button';btn.className='secondary';btn.id='exportPdf';btn.style.cssText='font-size:12px;padding:9px 12px';btn.textContent='↓ Boletas PDF';
 const wrap=document.createElement('div');wrap.style.cssText='display:flex;gap:8px;flex-wrap:wrap;justify-content:flex-end';
 const excelBtn=$('exportCsv');excelBtn.replaceWith(wrap);wrap.append(excelBtn,btn);
 let working=false;
 const originalRender=render;render=function(){originalRender();btn.disabled=working||!state.receipts.some(r=>r.photoUrl);};
 const clp=n=>'$'+Number(n).toLocaleString('es-CL');
 const dmy=d=>String(d).split('-').reverse().join('-');
 // Reduce cada foto para que el PDF no pese demasiado y obtiene sus dimensiones.
 function shrink(base64){return new Promise((resolve,reject)=>{const img=new Image();img.onload=()=>{const scale=Math.min(1,900/img.naturalWidth,2400/img.naturalHeight),w=Math.max(1,Math.round(img.naturalWidth*scale)),h=Math.max(1,Math.round(img.naturalHeight*scale));const c=document.createElement('canvas');c.width=w;c.height=h;const ctx=c.getContext('2d');ctx.fillStyle='#fff';ctx.fillRect(0,0,w,h);ctx.drawImage(img,0,0,w,h);resolve({data:c.toDataURL('image/jpeg',.85),w,h});};img.onerror=()=>reject(Error('No se pudo leer una foto.'));img.src='data:image/jpeg;base64,'+base64;});}
 btn.onclick=async()=>{
  if(working)return;if(!cloudReady||saving){toast('Espera a que termine la conexión o el guardado.');return;}
  const person=state.person,list=state.receipts.filter(r=>r.photoUrl),session=sessionEpoch;
  if(!list.length){toast('No hay boletas con foto para descargar.');return;}
  working=true;btn.disabled=true;const items=[];let failed=0;
  try{
   await loadScript('https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js',()=>!!(window.jspdf&&window.jspdf.jsPDF));
   for(let i=0;i<list.length;i++){
    btn.textContent=`Cargando fotos ${i+1}/${list.length}…`;
    try{const photo=await rpc('getPhoto',{id:list[i].id});items.push({r:list[i],img:await shrink(photo.data)});}
    catch(e){if(session!==sessionEpoch)return;failed++;}
   }
   if(session!==sessionEpoch)return;
   if(!items.length)throw Error('No se pudo cargar ninguna foto.');
   btn.textContent='Armando PDF…';
   const doc=new window.jspdf.jsPDF({unit:'mm',format:'letter',orientation:'portrait'});
   const W=215.9,H=279.4,M=10,gap=6,colW=(W-2*M-2*gap)/3,top=20,capH=16,bottom=H-M;
   const who=person==='__all__'?'Todas las personas':person,pages=Math.ceil(items.length/3);
   items.forEach((it,i)=>{
    const col=i%3;if(i&&!col)doc.addPage('letter','portrait');
    if(!col){
     doc.setFont('helvetica','bold');doc.setFontSize(11);doc.text('Boletas · '+who,M,M+4);
     doc.setFont('helvetica','normal');doc.setFontSize(8);doc.text(`Hoja ${Math.floor(i/3)+1} de ${pages}`,W-M,M+4,{align:'right'});
     doc.setDrawColor(190);doc.setLineDashPattern([1.5,1.5],0);for(const k of [1,2]){const x=M+k*colW+(k-.5)*gap;doc.line(x,top-2,x,bottom);}doc.setLineDashPattern([],0);
    }
    const x=M+col*(colW+gap),r=it.r;
    doc.setFont('helvetica','bold');doc.setFontSize(9);doc.text(`${i+1}. N.º ${r.number} · ${clp(r.amount)}`,x,top+3,{maxWidth:colW});
    doc.setFont('helvetica','normal');doc.setFontSize(8);
    const sub=dmy(r.date)+(person==='__all__'?' · '+r.person:'')+' · '+r.merchant;doc.text(doc.splitTextToSize(sub,colW).slice(0,2),x,top+7.5);
    const boxH=bottom-(top+capH),s=Math.min(colW/it.img.w,boxH/it.img.h),w=it.img.w*s,h=it.img.h*s;
    doc.addImage(it.img.data,'JPEG',x+(colW-w)/2,top+capH,w,h,undefined,'FAST');
   });
   doc.save(`boletas-${who.replace(/[^\p{L}\d]+/gu,'-')}-${today()}.pdf`);
   const skipped=state.receipts.length-list.length;
   toast(`PDF listo: ${items.length} boletas en ${pages} ${pages===1?'hoja':'hojas'}.`+(failed?(failed===1?' 1 foto no se pudo cargar.':` ${failed} fotos no se pudieron cargar.`):'')+(skipped?(skipped===1?' 1 boleta sin foto no se incluyó.':` ${skipped} boletas sin foto no se incluyeron.`):''));
  }catch(e){if(session===sessionEpoch)showError(e);}
  finally{working=false;btn.textContent='↓ Boletas PDF';btn.disabled=!state.receipts.some(r=>r.photoUrl);}
 };
 btn.disabled=!state.receipts.some(r=>r.photoUrl);
})();
