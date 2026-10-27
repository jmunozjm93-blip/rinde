/* Exportación en el formato de rendiciones de Comercial Depor. */
(function(root){
'use strict';
const categories=[
 ['Gastos generales','(Colación y Movilización)','61-08-001-01'],
 ['Artículos','de Oficina','61-08-001-08'],
 ['Peaje, Estacionamiento','y Combustible','61-08-001-13'],
 ['Insumos','Computación','61-07-001-03'],
 ['Alojamiento, Estadía y','Pasajes','61-08-001-03'],
 ['Otros','Honorarios',''],['Factura','de Gasto','']
];
const keys=["Gastos generales","Artículos de oficina","Peaje, estacionamiento y combustible","Insumos computación","Alojamiento, estadía y pasajes","Otros honorarios","Factura de gasto"];
function suggest(r){
 const saved=keys.indexOf(r.category);if(saved>=0)return saved;
 const t=(r.merchant||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase();
 if(r.category==='Alimentación')return 0;
 if(r.category==='Alojamiento')return 4;
 if(r.category==='Transporte'){
  if(/\b(peaje|tag|bencina|combustible|gasolina|diesel|estacionamiento|copec|shell|enex)\b/.test(t))return 2;
  if(/\b(avion|vuelo|aereo|aerolinea|latam|jetsmart|sky|pasaje)\b/.test(t))return 4;
  if(/\b(taxi|uber|cabify|didi|metro|micro|locomocion|movilizacion)\b/.test(t))return 0;
 }
 return null;
}
async function build(ExcelJS,data,meta,choices){
 if(!data.person||data.person==='__all__')throw Error('Selecciona una persona para descargar su rendición.');
 if(!data.receipts.length)throw Error('No hay boletas para descargar.');
 if(choices.length!==data.receipts.length||choices.some(c=>!Number.isInteger(c)||c<0||c>6))throw Error('Elige la columna de cada gasto.');
 if(!Number.isSafeInteger(data.budget)||data.budget<0)throw Error('Fondo no válido.');
 const wb=new ExcelJS.Workbook();wb.creator='Rinde';wb.calcProperties.fullCalcOnLoad=true;
 const ws=wb.addWorksheet('Rendición',{pageSetup:{paperSize:9,orientation:'landscape',fitToPage:true,fitToWidth:1,fitToHeight:0,margins:{left:.25,right:.25,top:.35,bottom:.35,header:.1,footer:.1}},views:[{state:'frozen',ySplit:12}]});
 [2,5,16,13,38,17,16,20,16,20,16,15,18,13].forEach((width,i)=>ws.getColumn(i+1).width=width);
 const last=12+Math.max(30,data.receipts.length),totalRow=last+1,signatureRow=last+5;
 ws.pageSetup.printArea=`B2:N${signatureRow+2}`;ws.pageSetup.printTitlesRow='10:12';
 const black={argb:'FF222222'},thin={style:'thin',color:black};
 function cell(a,v){const c=ws.getCell(a);c.value=v;return c;}
 function merge(a,v){ws.mergeCells(a);cell(a.split(':')[0],v);}
 for(let row=2;row<=signatureRow+2;row++){
  ws.getRow(row).height=22;
  for(let col=2;col<=14;col++){const c=ws.getCell(row,col);c.font={name:'Arial',size:10};c.alignment={vertical:'middle',wrapText:true};}
 }
 merge('E4:I4','COMERCIAL DEPOR LTDA.');ws.getCell('E4').font={name:'Arial',size:15,bold:true};ws.getCell('E4').alignment={horizontal:'center',vertical:'middle'};
 const heads=[['J2','N.º de Rendición:','number'],['J3','Periodo:','period'],['J4','Centro de costo:','costCenter'],['J5','Marca:','brand'],['J6','Canal:','channel'],['J7','Sucursal:','branch'],['J8','Código SAP:','sap']];
 for(const [a,label,key]of heads){cell(a,label).font={name:'Arial',size:10,bold:true};merge(`K${a.slice(1)}:M${a.slice(1)}`,String(meta[key]||''));}
 merge('B5:D5','Marque con una cruz:');cell('B6',meta.type==='Reembolso'?'X':'');merge('C6:D6','Reembolso');cell('B8',meta.type==='Fondo por rendir'?'X':'');cell('C8','Fondo x rendir');
 merge('D8:E8','Cheque N.º: '+(meta.check||''));merge('F8:G8','Banco: '+(meta.bank||''));cell('H8','Monto $');cell('I8',data.budget).numFmt='"$"#,##0';
 merge('B9:I9','Rendido por: '+data.person);if(data.paymentId)merge('J9:N9','Pago: '+data.paymentId);
 ['N.º','N.º Documento','Fecha','Descripción de Gasto'].forEach((v,i)=>merge(`${String.fromCharCode(66+i)}10:${String.fromCharCode(66+i)}12`,v));
 categories.forEach((cat,i)=>cat.forEach((v,j)=>cell(`${String.fromCharCode(70+i)}${10+j}`,v)));
 merge('M10:M12','Monto $');merge('N10:N12','Estado');
 for(let r=10;r<=12;r++)for(let c=2;c<=14;c++){const x=ws.getCell(r,c);x.font={name:'Arial',size:r===12?8:9,bold:true};x.alignment={horizontal:'center',vertical:'middle',wrapText:true};x.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFE9E9E9'}};x.border={top:thin,left:thin,right:thin,bottom:thin};}
 ws.getRow(10).height=30;ws.getRow(11).height=28;
 const totals=Array(7).fill(0);
 for(let row=13;row<=last;row++){
  const r=data.receipts[row-13];ws.getRow(row).height=r?Math.max(30,Math.ceil(r.merchant.length/32)*14):22;
  for(let col=2;col<=14;col++){const c=ws.getCell(row,col);c.border={top:thin,left:thin,right:thin,bottom:thin};if(col>=6)c.numFmt='#,##0;[Red](#,##0);"-"';}
  if(!r)continue;
  if(!Number.isSafeInteger(r.amount)||r.amount<=0)throw Error('Hay un monto no válido.');
  cell('N'+row,data.status==='Pagado'?'Pagado':'Pendiente');cell('B'+row,row-12);cell('C'+row,String(r.number)).numFmt='@';
  cell('D'+row,new Date(r.date+'T00:00:00Z')).numFmt='dd-mm-yyyy';cell('E'+row,String(r.merchant));
  const col=choices[row-13];cell(String.fromCharCode(70+col)+row,r.amount);totals[col]+=r.amount;
  cell('M'+row,{formula:`SUM(F${row}:L${row})`,result:r.amount});
 }
 merge(`B${totalRow}:E${totalRow}`,'TOTAL RENDIDO');
 totals.forEach((n,i)=>cell(`${String.fromCharCode(70+i)}${totalRow}`,{formula:`SUM(${String.fromCharCode(70+i)}13:${String.fromCharCode(70+i)}${last})`,result:n}));
 const total=totals.reduce((a,b)=>a+b,0);cell('M'+totalRow,{formula:`SUM(M13:M${last})`,result:total});
 for(let c=2;c<=14;c++){const x=ws.getCell(totalRow,c);x.font={name:'Arial',size:10,bold:true};x.border={top:thin,bottom:thin};x.numFmt='#,##0';}
 merge(`J${totalRow+1}:L${totalRow+1}`,'Saldo / pendiente de cubrir');cell('M'+(totalRow+1),{formula:`I8-M${totalRow}`,result:data.budget-total}).numFmt='#,##0;[Red](#,##0)';
 merge(`B${signatureRow}:E${signatureRow}`,data.person);merge(`G${signatureRow}:I${signatureRow}`,String(meta.manager||''));merge(`K${signatureRow}:M${signatureRow}`,'');
 merge(`B${signatureRow+1}:E${signatureRow+1}`,'Rendido Por');merge(`G${signatureRow+1}:I${signatureRow+1}`,'Jefe Directo');merge(`K${signatureRow+1}:M${signatureRow+1}`,'V.º B.º Auditor Interno');
 for(const a of [`B${signatureRow}`,`G${signatureRow}`,`K${signatureRow}`])ws.getCell(a).border={bottom:thin};
 ws.headerFooter.oddFooter='&CPágina &P de &N';
 return wb;
}
const api={categories,suggest,build};if(typeof module!=='undefined')module.exports=api;else root.RindeExcel=api;
})(typeof window==='undefined'?this:window);
