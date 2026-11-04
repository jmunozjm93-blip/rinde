'use strict';
/* Fotos de boletas en PDF tamaño carta, tres boletas por hoja (una por columna).
   Se usa desde la descarga de la rendición (export-ui.js), junto con el Excel. */
(()=>{
 const clp=n=>'$'+Number(n).toLocaleString('es-CL');
 const dmy=d=>String(d).split('-').reverse().join('-');
 // Reduce cada foto para que el PDF no pese demasiado y obtiene sus dimensiones.
 function shrink(base64){return new Promise((resolve,reject)=>{const img=new Image();img.onload=()=>{const scale=Math.min(1,900/img.naturalWidth,2400/img.naturalHeight),w=Math.max(1,Math.round(img.naturalWidth*scale)),h=Math.max(1,Math.round(img.naturalHeight*scale));const c=document.createElement('canvas');c.width=w;c.height=h;const ctx=c.getContext('2d');ctx.fillStyle='#fff';ctx.fillRect(0,0,w,h);ctx.drawImage(img,0,0,w,h);resolve({data:c.toDataURL('image/jpeg',.85),w,h});};img.onerror=()=>reject(Error('No se pudo leer una foto.'));img.src='data:image/jpeg;base64,'+base64;});}
 /* receipts: en el mismo orden que el Excel, así el número de cada boleta coincide con la columna N.º.
    Devuelve {blob,count,pages,failed,skipped}; blob es null si no hay ninguna foto. */
 async function build({receipts,title,showPerson,progress}){
  const session=sessionEpoch,list=receipts.map((r,i)=>({r,n:i+1})).filter(x=>x.r.photoUrl);
  const result={blob:null,count:0,pages:0,failed:0,skipped:receipts.length-list.length};
  if(!list.length)return result;
  await loadScript('https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js',()=>!!(window.jspdf&&window.jspdf.jsPDF));
  const items=[];
  for(let i=0;i<list.length;i++){
   if(progress)progress(`Cargando fotos de boletas ${i+1}/${list.length}…`);
   try{const photo=await rpc('getPhoto',{id:list[i].r.id});items.push({...list[i],img:await shrink(photo.data)});}
   catch(e){if(session!==sessionEpoch)throw Error('La sesión cambió.');result.failed++;}
  }
  if(session!==sessionEpoch)throw Error('La sesión cambió.');
  if(!items.length)return result;
  if(progress)progress('Armando PDF de boletas…');
  const doc=new window.jspdf.jsPDF({unit:'mm',format:'letter',orientation:'portrait'});
  const W=215.9,H=279.4,M=10,gap=6,colW=(W-2*M-2*gap)/3,top=20,capH=16,bottom=H-M,pages=Math.ceil(items.length/3);
  items.forEach((it,i)=>{
   const col=i%3;if(i&&!col)doc.addPage('letter','portrait');
   if(!col){
    doc.setFont('helvetica','bold');doc.setFontSize(11);doc.text(doc.splitTextToSize(title,W-2*M-30)[0],M,M+4);
    doc.setFont('helvetica','normal');doc.setFontSize(8);doc.text(`Hoja ${Math.floor(i/3)+1} de ${pages}`,W-M,M+4,{align:'right'});
    doc.setDrawColor(190);doc.setLineDashPattern([1.5,1.5],0);for(const k of [1,2]){const x=M+k*colW+(k-.5)*gap;doc.line(x,top-2,x,bottom);}doc.setLineDashPattern([],0);
   }
   const x=M+col*(colW+gap),r=it.r;
   doc.setFont('helvetica','bold');doc.setFontSize(9);doc.text(`${it.n}. N.º ${r.number} · ${clp(r.amount)}`,x,top+3,{maxWidth:colW});
   doc.setFont('helvetica','normal');doc.setFontSize(8);
   const sub=dmy(r.date)+(showPerson?' · '+r.person:'')+' · '+r.merchant;doc.text(doc.splitTextToSize(sub,colW).slice(0,2),x,top+7.5);
   const boxH=bottom-(top+capH),s=Math.min(colW/it.img.w,boxH/it.img.h),w=it.img.w*s,h=it.img.h*s;
   doc.addImage(it.img.data,'JPEG',x+(colW-w)/2,top+capH,w,h,undefined,'FAST');
  });
  return {...result,blob:doc.output('blob'),count:items.length,pages};
 }
 function summary(res){
  if(!res.blob)return res.failed?'No se pudieron cargar las fotos de las boletas; se descarga solo el Excel.':'Estas boletas no tienen fotos; se descarga solo el Excel.';
  return `PDF de boletas: ${res.count} en ${res.pages} ${res.pages===1?'hoja':'hojas'}.`+(res.failed?(res.failed===1?' 1 foto no se pudo cargar.':` ${res.failed} fotos no se pudieron cargar.`):'')+(res.skipped?(res.skipped===1?' 1 boleta sin foto no se incluyó.':` ${res.skipped} boletas sin foto no se incluyeron.`):'');
 }
 const slug=s=>String(s).replace(/[^\p{L}\d]+/gu,'-').replace(/^-|-$/g,'');
 window.RindePdf={build,summary,slug};
})();
