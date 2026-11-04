'use strict';
/* PDF tamaño carta de la rendición: primero la planilla con el mismo formato del Excel de Comercial Depor
   (hoja horizontal) y después las fotos de las boletas: las chicas de a tres por hoja, las grandes solas en su hoja.
   Se usa desde la descarga de la rendición (export-ui.js). */
(()=>{
 const clp=n=>'$'+Number(n).toLocaleString('es-CL');
 const dmy=d=>String(d).split('-').reverse().join('-');
 const num=n=>!n?'-':n<0?'('+Math.abs(n).toLocaleString('es-CL')+')':Number(n).toLocaleString('es-CL');
 const M=10;
 // Reduce cada foto para que el PDF no pese demasiado y obtiene sus dimensiones.
 function shrink(base64){return new Promise((resolve,reject)=>{const img=new Image();img.onload=()=>{const scale=Math.min(1,1600/img.naturalWidth,2400/img.naturalHeight),w=Math.max(1,Math.round(img.naturalWidth*scale)),h=Math.max(1,Math.round(img.naturalHeight*scale));const c=document.createElement('canvas');c.width=w;c.height=h;const ctx=c.getContext('2d');ctx.fillStyle='#fff';ctx.fillRect(0,0,w,h);ctx.drawImage(img,0,0,w,h);resolve({data:c.toDataURL('image/jpeg',.85),w,h});};img.onerror=()=>reject(Error('No se pudo leer una foto.'));img.src='data:image/jpeg;base64,'+base64;});}

 /* Planilla de la rendición con el formato del Excel: encabezado solo con los campos que tienen dato,
    tabla por columna contable, total, saldo y firmas. */
 function drawSheet(doc,{data,meta,choices}){
  const W=279.4,H=215.9,val=k=>String(meta[k]||'').trim();
  doc.setDrawColor(34);doc.setLineWidth(.25);doc.setTextColor(20);
  doc.setFont('helvetica','bold');doc.setFontSize(15);doc.text('COMERCIAL DEPOR LTDA.',W/2-25,24,{align:'center'});
  const heads=[['N.º de Rendición:','number'],['Periodo:','period'],['Centro de costo:','costCenter'],['Marca:','brand'],['Canal:','channel'],['Sucursal:','branch'],['Código SAP:','sap']].filter(([,k])=>val(k));
  const bw=55,bx=W-M-bw,lx=bx-31;
  heads.forEach(([label,k],i)=>{const y=12+i*6;doc.setFont('helvetica','bold');doc.setFontSize(9);doc.text(label,lx,y+4.2);doc.rect(bx,y,bw,6);doc.setFont('helvetica','normal');doc.text(doc.splitTextToSize(val(k),bw-3)[0],bx+1.5,y+4.2);});
  doc.setFont('helvetica','bold');doc.setFontSize(9);const mark='Marque con una cruz:';doc.text(mark,M,34);doc.line(M,35,M+doc.getTextWidth(mark),35);
  const box=(y,label,on)=>{doc.rect(M,y,5,5);if(on)doc.text('X',M+2.5,y+3.8,{align:'center'});doc.text(label,M+7,y+3.8);};
  box(38,'Reembolso',meta.type==='Reembolso');box(46,'Fondo x rendir',meta.type==='Fondo por rendir');
  let x=M+42;
  for(const t of [val('check')&&'Cheque N.º: '+val('check'),val('bank')&&'Banco: '+val('bank'),data.budget>0&&'Monto $ '+num(data.budget)])if(t){doc.text(t,x,49.8);x+=doc.getTextWidth(t)+12;}
  const listTop=Math.max(60,12+heads.length*6+6);
  doc.text('Rendido por: '+data.person,M,listTop);
  if(data.paymentId){doc.setFont('helvetica','normal');doc.setFontSize(7);doc.text('Pago: '+data.paymentId,W-M,listTop,{align:'right'});}
  const cats=RindeExcel.categories,totals=Array(7).fill(0);
  const body=data.receipts.map((r,i)=>{const row=Array(7).fill('');row[choices[i]]=num(r.amount);totals[choices[i]]+=r.amount;return [String(i+1),String(r.number),dmy(r.date),String(r.merchant),...row,num(r.amount),data.status==='Pagado'?'Pagado':'Pendiente'];});
  const total=totals.reduce((a,b)=>a+b,0),catStyle={cellWidth:19,halign:'right'},columnStyles={0:{cellWidth:8,halign:'center'},1:{cellWidth:20},2:{cellWidth:18},3:{cellWidth:'auto'},11:{cellWidth:20,halign:'right'},12:{cellWidth:16,halign:'center'}};
  for(let c=4;c<=10;c++)columnStyles[c]=catStyle;
  doc.autoTable({startY:listTop+4,margin:{left:M,right:M,top:15,bottom:15},theme:'grid',showFoot:'lastPage',
   head:[['N.º','N.º Documento','Fecha','Descripción de Gasto',...cats.map(c=>c[0]+' '+c[1]+(c[2]?'\n'+c[2]:'')),'Monto $','Estado']],
   body,foot:[[{content:'TOTAL RENDIDO',colSpan:4},...totals.map(num),num(total),'']],
   styles:{font:'helvetica',fontSize:7.5,cellPadding:1.2,lineColor:[34,34,34],lineWidth:.2,textColor:20,valign:'middle',overflow:'linebreak'},
   headStyles:{fillColor:[233,233,233],textColor:20,fontStyle:'bold',halign:'center',fontSize:6.5},
   footStyles:{fillColor:[255,255,255],textColor:20,fontStyle:'bold',halign:'right'},columnStyles});
  let y=doc.lastAutoTable.finalY+8;if(y>H-45){doc.addPage('letter','landscape');y=20;}
  const saldo=data.budget-total;doc.setFont('helvetica','bold');doc.setFontSize(9);doc.text('Saldo / pendiente de cubrir:',W-M-40,y,{align:'right'});
  if(saldo<0)doc.setTextColor(200,0,0);doc.text(num(saldo),W-M-2,y,{align:'right'});doc.setTextColor(20);
  const sy=y+24,sw=70,sx=[M+5,(W-sw)/2,W-M-5-sw],names=[data.person,val('manager'),''],labels=['Rendido Por','Jefe Directo','V.º B.º Auditor Interno'];
  sx.forEach((x0,i)=>{doc.setFont('helvetica','normal');doc.setFontSize(9);if(names[i])doc.text(names[i],x0+sw/2,sy-2,{align:'center'});doc.line(x0,sy,x0+sw,sy);doc.setFont('helvetica','bold');doc.text(labels[i],x0+sw/2,sy+5,{align:'center'});});
 }

 /* receipts: en el mismo orden que el Excel, así el número de cada boleta coincide con la columna N.º.
    sheet (opcional): {data,meta,choices} para incluir primero la planilla de la rendición.
    Devuelve {blob,count,pages,failed,skipped,sheet}; blob es null solo si no hay planilla ni fotos. */
 async function build({receipts,title,showPerson,progress,sheet}){
  const session=sessionEpoch,list=receipts.map((r,i)=>({r,n:i+1})).filter(x=>x.r.photoUrl);
  const result={blob:null,count:0,pages:0,failed:0,skipped:receipts.length-list.length,sheet:!!sheet};
  if(!list.length&&!sheet)return result;
  await loadScript('https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js',()=>!!(window.jspdf&&window.jspdf.jsPDF));
  if(sheet)await loadScript('https://cdn.jsdelivr.net/npm/jspdf-autotable@3.8.2/dist/jspdf.plugin.autotable.min.js',()=>!!(window.jspdf.jsPDF.API.autoTable));
  const items=[];
  for(let i=0;i<list.length;i++){
   if(progress)progress(`Cargando fotos de boletas ${i+1}/${list.length}…`);
   try{const photo=await rpc('getPhoto',{id:list[i].r.id});items.push({...list[i],img:await shrink(photo.data)});}
   catch(e){if(session!==sessionEpoch)throw Error('La sesión cambió.');result.failed++;}
  }
  if(session!==sessionEpoch)throw Error('La sesión cambió.');
  if(!items.length&&!sheet)return result;
  if(progress)progress('Armando PDF…');
  /* Boletas chicas (angostas y largas, como las térmicas) se agrupan de a 3 por hoja, una por columna.
     Boletas grandes (hoja completa, foto ancha o casi cuadrada) van solas en su hoja, ocupando toda la página;
     si la foto es horizontal, la hoja va horizontal. */
  const isSmall=it=>it.img.h/it.img.w>=1.6;
  const plan=[],small=items.filter(isSmall);
  for(let i=0;i<small.length;i+=3)plan.push({group:true,items:small.slice(i,i+3),landscape:false});
  for(const it of items.filter(it=>!isSmall(it)))plan.push({group:false,items:[it],landscape:it.img.w>it.img.h});
  const gap=6,top=20,capH=16;let doc=null;
  const newPage=landscape=>{const o=landscape?'landscape':'portrait';if(!doc)doc=new window.jspdf.jsPDF({unit:'mm',format:'letter',orientation:o});else doc.addPage('letter',o);};
  if(sheet){newPage(true);drawSheet(doc,sheet);}
  function caption(it,x,width){
   const r=it.r;doc.setTextColor(20);doc.setFont('helvetica','bold');doc.setFontSize(9);doc.text(`${it.n}. N.º ${r.number} · ${clp(r.amount)}`,x,top+3,{maxWidth:width});
   doc.setFont('helvetica','normal');doc.setFontSize(8);
   const sub=dmy(r.date)+(showPerson?' · '+r.person:'')+' · '+r.merchant;doc.text(doc.splitTextToSize(sub,width).slice(0,2),x,top+7.5);
  }
  function image(it,x,y,boxW,boxH){const s=Math.min(boxW/it.img.w,boxH/it.img.h),w=it.img.w*s,h=it.img.h*s;doc.addImage(it.img.data,'JPEG',x+(boxW-w)/2,y,w,h,undefined,'FAST');}
  plan.forEach(page=>{
   newPage(page.landscape);
   const W=page.landscape?279.4:215.9,H=page.landscape?215.9:279.4,bottom=H-M-4;
   doc.setTextColor(20);doc.setFont('helvetica','bold');doc.setFontSize(11);doc.text(doc.splitTextToSize('Boletas · '+title,W-2*M)[0],M,M+4);
   if(page.group){
    const colW=(W-2*M-2*gap)/3;
    doc.setDrawColor(190);doc.setLineDashPattern([1.5,1.5],0);for(const k of [1,2]){const x=M+k*colW+(k-.5)*gap;doc.line(x,top-2,x,bottom);}doc.setLineDashPattern([],0);
    page.items.forEach((it,col)=>{const x=M+col*(colW+gap);caption(it,x,colW);image(it,x,top+capH,colW,bottom-(top+capH));});
   }else{const it=page.items[0];caption(it,M,W-2*M);image(it,M,top+capH,W-2*M,bottom-(top+capH));}
  });
  // Numeración de todas las hojas al pie.
  const pages=doc.getNumberOfPages();
  for(let p=1;p<=pages;p++){doc.setPage(p);const W=doc.internal.pageSize.getWidth(),H=doc.internal.pageSize.getHeight();doc.setTextColor(110);doc.setFont('helvetica','normal');doc.setFontSize(8);doc.text(`Hoja ${p} de ${pages}`,W/2,H-5,{align:'center'});}
  return {...result,blob:doc.output('blob'),count:items.length,pages};
 }
 function summary(res){
  const lost=(res.failed?(res.failed===1?' 1 foto no se pudo cargar.':` ${res.failed} fotos no se pudieron cargar.`):'')+(res.skipped?(res.skipped===1?' 1 boleta sin foto no se incluyó.':` ${res.skipped} boletas sin foto no se incluyeron.`):'');
  if(!res.blob)return res.failed?'No se pudieron cargar las fotos de las boletas; se descarga solo el Excel.':'Estas boletas no tienen fotos; se descarga solo el Excel.';
  const hojas=`${res.pages} ${res.pages===1?'hoja':'hojas'}`;
  if(res.sheet)return (res.count?`PDF listo: rendición + ${res.count} ${res.count===1?'boleta':'boletas'} en ${hojas}.`:`PDF listo con la rendición (${hojas}); no trae fotos de boletas.`)+lost;
  return `PDF de boletas: ${res.count} en ${hojas}.`+lost;
 }
 const slug=s=>String(s).replace(/[^\p{L}\d]+/gu,'-').replace(/^-|-$/g,'');
 window.RindePdf={build,summary,slug};
})();
