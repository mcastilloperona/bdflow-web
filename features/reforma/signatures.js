(function(root,factory){if(typeof module==='object'&&module.exports)module.exports=factory(require('./core.js'));else root.ReformaSignatures=factory(root.ReformaCore);})(typeof globalThis!=='undefined'?globalThis:this,function(C){
 'use strict';
 function rows(imports,snapshots,signatures,today){
  const dates=new Map(imports.map(i=>[i.id,i])),saved=new Map(signatures.map(v=>[v.operation_id,v])),latest=new Map(),months=new Map();
  const ordered=snapshots.filter(v=>dates.has(v.import_id)).slice().sort((a,b)=>C.importCompare(dates.get(a.import_id),dates.get(b.import_id)));
  for(const v of ordered){const r=v.data;if(C.inPlan(r,C.REFORMA)){if(!months.has(v.operation_id))months.set(v.operation_id,new Set());if(r.close_date)months.get(v.operation_id).add(r.close_date.slice(0,7));}latest.set(v.operation_id,{...r,operation_id:v.operation_id,report_date:dates.get(v.import_id).report_date});}
  return [...latest.values()].filter(r=>months.has(r.operation_id)).map(r=>{const signature=saved.get(r.operation_id);return {...r,months:[...months.get(r.operation_id)].sort(),signature,signature_status:signature?.status||'unconfirmed',overdue:signature?.status!=='signed'&&!!signature?.next_followup&&signature.next_followup<today};}).sort((a,b)=>Number(b.overdue)-Number(a.overdue)||(a.signature?.next_followup||'9999').localeCompare(b.signature?.next_followup||'9999')||String(a.account).localeCompare(String(b.account)));
 }
 function validate(v){if(!['pending','signed'].includes(v.status))throw Error('Selecciona el estado de firma.');if(v.status==='signed'&&!v.signed_date)throw Error('Indica la fecha de firma.');if(v.status==='pending'&&v.signed_date)throw Error('Una firma pendiente no puede tener fecha de firma.');if((v.notes||'').length>10000)throw Error('El comentario supera los 10000 caracteres.');return v;}
 return {rows,validate};
});
