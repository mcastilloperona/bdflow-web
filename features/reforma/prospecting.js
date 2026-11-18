(function(root,factory){if(typeof module==='object'&&module.exports)module.exports=factory();else root.ReformaProspecting=factory();})(typeof globalThis!=='undefined'?globalThis:this,function(){
 'use strict';
 function date(value){let v=String(value||'').trim(),m=v.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);if(m)v=m[3]+'-'+m[2].padStart(2,'0')+'-'+m[1].padStart(2,'0');if(!/^\d{4}-\d{2}-\d{2}$/.test(v)||!Number.isFinite(Date.parse(v+'T00:00:00Z'))||new Date(v+'T00:00:00Z').toISOString().slice(0,10)!==v)throw Error('Fecha no válida: '+String(value).slice(0,40));return v;}
 const norm=v=>String(v||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().trim();
 function parse(text,reportDate){
  const cutoff=date(reportDate),seen=new Set(),rows=[];if(String(text).length>100000)throw Error('El reporte de prospección supera el tamaño permitido');
  let names=null,hasTotal=false;
  const count=(raw,line)=>{if(!/^\d+$/.test(raw)||+raw>100000)throw Error('Línea '+line+': el recuento debe ser un entero entre 0 y 100 000');return +raw;};
  for(const [i,line] of String(text).split(/\r?\n/).entries()){
   if(!line.trim())continue;
   const cells=names?line.split('\t').map(v=>v.trim()):line.trim().split(/[\t;,]+|\s+/);
   if(/^fecha/i.test(cells[0])){const head=line.trim().split('\t').map(v=>v.trim());if(head.length>2){hasTotal=/^total(?: general)?$/i.test(head.at(-1));names=head.slice(1,hasTotal?-1:undefined);if(names.some(n=>!n||n.length>80)||new Set(names.map(norm)).size!==names.length)throw Error('Los nombres del reporte deben ser únicos');}continue;}
   const day=date(cells[0]);if(day>cutoff)throw Error('Una fecha de llamada es posterior a la fecha del reporte');if(seen.has(day))throw Error('Fecha duplicada en el reporte: '+day);seen.add(day);
   if(names){if(cells.length!==1+names.length+(hasTotal?1:0))throw Error('Línea '+(i+1)+': revisa las columnas del reporte');let total=0;names.forEach((caller,j)=>{if(cells[j+1]==='')return;const n=count(cells[j+1],i+1);total+=n;rows.push({date:day,count:n,caller});});if(hasTotal&&count(cells.at(-1),i+1)!==total)throw Error('Línea '+(i+1)+': el total no coincide con el desglose');}
   else{if(cells.length!==2)throw Error('Línea '+(i+1)+': introduce fecha y total o una tabla con encabezados por persona');rows.push({date:day,count:count(cells[1],i+1)});}
  }
  if(!rows.length)throw Error('Introduce al menos una fecha con su recuento');return rows.sort((a,b)=>a.date.localeCompare(b.date));
 }
 function serialize(rows){
  const names=[...new Set(rows.map(r=>r.caller).filter(Boolean))];if(!names.length)return rows.map(r=>r.date+'\t'+r.count).join('\n');
  const dates=[...new Set(rows.map(r=>r.date))].sort();return [['Fecha',...names].join('\t'),...dates.map(day=>[day,...names.map(name=>{const r=rows.find(r=>r.date===day&&r.caller===name);return r?r.count:'';})].join('\t'))].join('\n');
 }
 function model(reports,asof,first){const report=reports.filter(r=>r.report_date<=asof).slice().sort((a,b)=>b.report_date.localeCompare(a.report_date)||(b.created_at||'').localeCompare(a.created_at||'')||(b.id||'').localeCompare(a.id||''))[0];if(!report)return {available:false,months:[],people:[],total:null,asof:null};try{
  const rows=parse(report.entries_text,report.report_date).filter(r=>r.date>=first&&r.date<=asof),months=new Map(),people=new Map();for(const r of rows){const month=r.date.slice(0,7);months.set(month,(months.get(month)||0)+r.count);if(r.caller){if(!people.has(r.caller))people.set(r.caller,{name:r.caller,count:0,activeDays:0});const p=people.get(r.caller);p.count+=r.count;if(r.count>0)p.activeDays++;}}
  const total=rows.reduce((n,r)=>n+r.count,0);return {available:true,report,asof:report.report_date,rows,months:[...months].map(([month,count])=>({month,count})),people:[...people.values()].map(p=>({...p,share:total?p.count/total:null,perActiveDay:p.activeDays?p.count/p.activeDays:null})).sort((a,b)=>b.count-a.count||a.name.localeCompare(b.name)),total,complete:report.record_complete===true,source:report.source_label,notes:report.notes||''};
 }catch(e){return {available:false,months:[],people:[],total:null,asof:report.report_date,error:e.message};}}
 function conversion(person,owners=[]){const tokens=norm(person).split(/\s+/),matches=owners.filter(o=>tokens.every(t=>norm(o.owner).split(/\s+/).includes(t)));return matches.length===1?matches[0]:null;}
 return {date,parse,serialize,model,conversion};
});
