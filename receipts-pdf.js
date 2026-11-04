'use strict';
/* Fotos de boletas en PDF tamaño carta: las chicas de a tres por hoja, las grandes solas en su hoja.
   Se usa desde la descarga de la rendición (export-ui.js), junto con el Excel. */
(()=>{
 const clp=n=>'$'+Number(n).toLocaleString('es-CL');
 const dmy=d=>String(d).split('-').reverse().join('-');
 // Reduce cada foto para que el PDF no pese demasiado y obtiene sus dimensiones.
 function shrink(base64){return new Promise((resolve,reject)=>{const img=new Image();img.onload=()=>{const scale=Math.min(1,1600/img.naturalWidth,2400/img.naturalHeight),w=Math.max(1,Math.round(img.naturalWidth*scale)),h=Math.max(1,Math.round(img.naturalHeight*scale));const c=document.createElement('canvas');c.width=w;c.height=h;const ctx=c.getContext('2d');ctx.fillStyle='#fff';ctx.fillRect(0,0,w,h);ctx.drawImage(img,0,0,w,h);resolve({data:c.toDataURL('image/jpeg',.85),w,h});};img.onerror=()=>reject(Error('No se pudo leer una foto.'));img.src='data:image/jpeg;base64,'+base64;});}
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
  /* Boletas chicas (angostas y largas, como las térmicas) se agrupan de a 3 por hoja, una por columna.
     Boletas grandes (hoja completa, foto ancha o casi cuadrada) van solas en su hoja, ocupando toda la página;
     si la foto es horizontal, la hoja va horizontal. */
  const isSmall=it=>it.img.h/it.img.w>=1.6;
  const plan=[],small=items.filter(isSmall);
  for(let i=0;i<small.length;i+=3)plan.push({group:true,items:small.slice(i,i+3),landscape:false});
  for(const it of items.filter(it=>!isSmall(it)))plan.push({group:false,items:[it],landscape:it.img.w>it.img.h});
  const M=10,gap=6,top=20,capH=16,pages=plan.length;
  const doc=new window.jspdf.jsPDF({unit:'mm',format:'letter',orientation:plan[0].landscape?'landscape':'portrait'});
  function caption(it,x,width){
   const r=it.r;doc.setFont('helvetica','bold');doc.setFontSize(9);doc.text(`${it.n}. N.º ${r.number} · ${clp(r.amount)}`,x,top+3,{maxWidth:width});
   doc.setFont('helvetica','normal');doc.setFontSize(8);
   const sub=dmy(r.date)+(showPerson?' · '+r.person:'')+' · '+r.merchant;doc.text(doc.splitTextToSize(sub,width).slice(0,2),x,top+7.5);
  }
  function image(it,x,y,boxW,boxH){const s=Math.min(boxW/it.img.w,boxH/it.img.h),w=it.img.w*s,h=it.img.h*s;doc.addImage(it.img.data,'JPEG',x+(boxW-w)/2,y,w,h,undefined,'FAST');}
  plan.forEach((page,p)=>{
   if(p)doc.addPage('letter',page.landscape?'landscape':'portrait');
   const W=page.landscape?279.4:215.9,H=page.landscape?215.9:279.4,bottom=H-M;
   doc.setFont('helvetica','bold');doc.setFontSize(11);doc.text(doc.splitTextToSize(title,W-2*M-30)[0],M,M+4);
   doc.setFont('helvetica','normal');doc.setFontSize(8);doc.text(`Hoja ${p+1} de ${pages}`,W-M,M+4,{align:'right'});
   if(page.group){
    const colW=(W-2*M-2*gap)/3;
    doc.setDrawColor(190);doc.setLineDashPattern([1.5,1.5],0);for(const k of [1,2]){const x=M+k*colW+(k-.5)*gap;doc.line(x,top-2,x,bottom);}doc.setLineDashPattern([],0);
    page.items.forEach((it,col)=>{const x=M+col*(colW+gap);caption(it,x,colW);image(it,x,top+capH,colW,bottom-(top+capH));});
   }else{const it=page.items[0];caption(it,M,W-2*M);image(it,M,top+capH,W-2*M,bottom-(top+capH));}
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
