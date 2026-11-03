(function(root,factory){if(typeof module==='object'&&module.exports)module.exports=factory(require('./core.js'));else root.ReformaSignatures=factory(root.ReformaCore);})(typeof globalThis!=='undefined'?globalThis:this,function(C){
 'use strict';
 function rows(imports,snapshots,signatures,today){
  const dates=new Map(imports.map(i=>[i.id,i])),saved=new Map(signatures.map(v=>[v.operation_id,v])),latest=new Map(),months=new Map();
  const ordered=snapshots.filter(v=>dates.has(v.import_id)).slice().sort((a,b)=>C.importCompare(dates.get(a.import_id),dates.get(b.import_id)));
  for(const v of ordered){const r=v.data;if(C.inPlan(r,C.REFORMA)){if(!months.has(v.operation_id))months.set(v.operation_id,new Set());if(r.close_date)months.get(v.operation_id).add(r.close_date.slice(0,7));}latest.set(v.operation_id,{...r,operation_id:v.operation_id,report_date:dates.get(v.import_id).report_date});}
  return [...latest.values()].filter(r=>C.inPlan(r,C.REFORMA)&&r.stage==='Closed Won'&&r.close_date&&r.close_date>='2026-10-01').map(r=>{const signature=saved.get(r.operation_id);return {...r,months:[r.close_date.slice(0,7)],signature,signature_status:signature?.status||'unconfirmed',overdue:signature?.status!=='signed'&&!!signature?.next_followup&&signature.next_followup<today};}).sort((a,b)=>Number(b.overdue)-Number(a.overdue)||(a.signature?.next_followup||'9999').localeCompare(b.signature?.next_followup||'9999')||String(a.account).localeCompare(String(b.account)));
 }
 function validate(v){if(!['pending','signed'].includes(v.status))throw Error('Selecciona el estado de firma.');if(v.status==='signed'&&!v.signed_date)throw Error('Indica la fecha de firma.');if(v.status==='pending'&&v.signed_date)throw Error('Una firma pendiente no puede tener fecha de firma.');if((v.notes||'').length>10000)throw Error('El comentario supera los 10000 caracteres.');return v;}

 function filteredRows(all,status='pending',month='all'){return all.filter(r=>(month==='all'||r.close_date.slice(0,7)===month)&&(status==='all'||status==='pending'&&r.signature_status!=='signed'||status==='signed'&&r.signature_status==='signed'||status==='overdue'&&r.overdue));}
 async function workbook(rows,ExcelJS,logoBytes,filters){
  const book=new ExcelJS.Workbook();book.creator='BDFlow';book.company='Cegid';book.title='Seguimiento de firmas Reforma Laboral México';
  const sh=book.addWorksheet('Firmas');sh.columns=[{width:32},{width:52},{width:22},{width:20},{width:20},{width:21},{width:20},{width:22},{width:24},{width:60},{width:30},{width:20}];
  const navy='123E69',blue='2962FF';sh.views=[{showGridLines:false,state:'frozen',xSplit:1,ySplit:7}];
  if(logoBytes){const logo=book.addImage({buffer:logoBytes,extension:'png'});sh.addImage(logo,{tl:{col:0,row:0},ext:{width:145,height:60}});}
  sh.mergeCells('C2:L2');sh.getCell('C2').value='Seguimiento de firmas Reforma Laboral México';sh.getCell('C2').font={name:'Inter',size:16,bold:true,color:{argb:navy}};
  sh.mergeCells('C3:L3');sh.getCell('C3').value='Generado el '+filters.today.split('-').reverse().join('/')+'   Estado: '+filters.statusLabel+'   Mes de cierre: '+filters.monthLabel;
  sh.mergeCells('A5:L5');sh.getCell('A5').value='Solo Closed Won con fecha de cierre Salesforce desde el 01/10/2026. Último dato disponible por oportunidad.';
  sh.mergeCells('A6:L6');sh.getCell('A6').value=rows.length+' oportunidades en la selección';
  for(const cell of ['C3','A5','A6']){sh.getCell(cell).font={name:'Inter',size:11,color:{argb:navy}};sh.getCell(cell).alignment={wrapText:true,vertical:'middle'};}
  sh.getRow(5).height=30;sh.getRow(7).values=['Cuenta','Oportunidad','Etapa Salesforce','Cierre Salesforce','Estado de firma','Fecha de firma','Próximo seguimiento','Seguimiento vencido','SaaS Booking (EUR)','Comentarios','Propietario Salesforce','Última aparición'];sh.getRow(7).height=32;
  sh.getRow(7).eachCell(c=>{c.font={name:'Inter',size:11,bold:true,color:{argb:'FFFFFF'}};c.fill={type:'pattern',pattern:'solid',fgColor:{argb:navy}};c.alignment={wrapText:true,vertical:'middle'};});
  const date=v=>v?new Date(v+'T00:00:00Z'):null;
  rows.forEach((r,i)=>{const v=r.signature||{},row=sh.getRow(i+8);row.values=[r.account,r.opportunity,r.stage,date(r.close_date),r.signature_status==='signed'?'Firmado':r.signature_status==='pending'?'Pendiente':'Sin confirmar',date(v.signed_date),date(v.next_followup),r.overdue?'Sí':'No',C.saas(r),v.notes||'',r.owner||'',date(r.report_date)];
   row.height=Math.max(38,Math.min(180,Math.ceil((v.notes||'').length/55)*15));row.eachCell(c=>{c.font={name:'Inter',size:11,color:{argb:navy}};c.alignment={wrapText:true,vertical:'top'};if(i%2===0)c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'F3F6FB'}};});
   for(const col of [4,6,7,12])row.getCell(col).numFmt='dd/mm/yyyy';row.getCell(9).numFmt='#,##0.00 "€"';
   if(r.overdue)for(const col of [7,8]){row.getCell(col).font={name:'Inter',size:11,bold:true,color:{argb:'B91C1C'}};row.getCell(col).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFF1F2'}};}
   row.getCell(5).font={name:'Inter',size:11,bold:true,color:{argb:blue}};
  });
  sh.autoFilter={from:{row:7,column:1},to:{row:Math.max(7,7+rows.length),column:12}};
  sh.pageSetup={orientation:'landscape',paperSize:9,fitToPage:true,fitToWidth:1,fitToHeight:0,printTitlesRow:'1:7'};
  return book.xlsx.writeBuffer();
 }
 return {rows,validate,filteredRows,workbook};
});
