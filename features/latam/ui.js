'use strict';
const latamReportUI={from:'',to:'',busy:false,error:''};
function currentLatamReport(){return BDLatamReport.model(getLatamActions(),state.tasks,latamReportUI);}
function setLatamReportFilter(key,value){if(latamReportUI.busy||!['from','to'].includes(key))return;latamReportUI[key]=value;latamReportUI.error='';render();}
function setLatamReportPeriod(period){
  if(latamReportUI.busy)return;
  const now=new Date(),local=d=>[d.getFullYear(),String(d.getMonth()+1).padStart(2,'0'),String(d.getDate()).padStart(2,'0')].join('-');
  latamReportUI.to=period==='all'?'':local(now);
  if(period==='month')latamReportUI.from=local(new Date(now.getFullYear(),now.getMonth(),1));
  else if(period==='week'){const monday=new Date(now);monday.setDate(now.getDate()-((now.getDay()+6)%7));latamReportUI.from=local(monday);}
  else latamReportUI.from='';
  latamReportUI.error='';render();
}
function renderLatamExecutiveReport(){
  let report=null,error=latamReportUI.error;try{report=currentLatamReport();}catch(e){error=e.message;}
  const disabled=latamReportUI.busy?'disabled':'';
  const dates=`<div class="latam-report-filters"><label>Desde<input type="date" value="${escapeHtml(latamReportUI.from)}" onchange="setLatamReportFilter('from',this.value)" ${disabled}></label><label>Hasta<input type="date" value="${escapeHtml(latamReportUI.to)}" onchange="setLatamReportFilter('to',this.value)" ${disabled}></label><div class="latam-report-shortcuts"><button class="btn secondary small" onclick="setLatamReportPeriod('week')" ${disabled}>Esta semana</button><button class="btn secondary small" onclick="setLatamReportPeriod('month')" ${disabled}>Este mes</button><button class="btn secondary small" onclick="setLatamReportPeriod('all')" ${disabled}>Todo el histórico</button></div></div>`;
  const header=`<div class="panel-head"><div><h2>Informe ejecutivo LATAM</h2><p>Actividad realizada, comentarios o resultados y próximos pasos para Dirección.</p></div><button class="btn primary" onclick="exportLatamPresentation()" ${disabled||(!report?.count?'disabled':'')}>${latamReportUI.busy?'Generando PowerPoint…':'Descargar PowerPoint'}</button></div>`;
  if(!report||!report.count)return `<div class="panel latam-executive-report">${header}<div class="panel-body">${dates}${error?`<p role="alert">${escapeHtml(error)}</p>`:''}<div class="empty">${report?'No hay colaboraciones en el periodo seleccionado.':'Revisa las fechas del informe.'}</div></div></div>`;
  const n=BDLatamReport.number,fd=BDLatamReport.fmtDate;
  return `<div class="panel latam-executive-report">${header}<div class="panel-body">${dates}${error?`<p role="alert">${escapeHtml(error)}</p>`:''}<p class="latam-report-period">${escapeHtml(report.period)}</p><div class="metrics"><div class="metric"><div class="k">ACTUACIONES</div><div class="v">${report.count}</div><div class="s">Actividad del periodo</div></div><div class="metric"><div class="k">COLABORACIONES</div><div class="v">${report.groups.length}</div><div class="s">Personas, equipos o áreas</div></div><div class="metric"><div class="k">TIEMPO DE APOYO</div><div class="v">${n(report.totalHours)} h</div><div class="s">Dedicación registrada</div></div></div><p class="latam-report-note">Actuaciones con comentario o resultado: ${report.withContext} de ${report.count}. Tareas BDFlow vinculadas: ${report.linkedCount}. Abiertas: ${report.next.length}.</p><h3>Colaboración por persona o equipo</h3><div class="latam-report-groups">${report.groups.map(g=>`<div><strong>${escapeHtml(g.name)}</strong><span>Actuaciones: ${g.count}. Apoyo: ${n(g.hours)} h.</span></div>`).join('')}</div><h3>Actuaciones y resultados registrados</h3>${report.rows.map(a=>`<article class="latam-report-action"><p class="latam-report-meta">${fd(a.date)}${a.time?' '+escapeHtml(a.time):''} <span>${escapeHtml(a.collaboration)}</span></p><h4>${escapeHtml(a.task)}</h4><p><strong>Comentario o resultado</strong><br>${escapeHtml(a.context||'Sin registrar')}</p><small>Tiempo de apoyo: ${n(a.hours)} h</small></article>`).join('')}<h3>Próximos pasos de tareas vinculadas</h3><p class="latam-report-note">Estado actual de las tareas vinculadas a la actividad del periodo.</p>${report.next.length?report.next.map(t=>`<article class="latam-report-action"><h4>${escapeHtml(t.title)}</h4><p>${escapeHtml(t.next||'Próximo paso sin registrar')}</p><p class="latam-report-meta">${escapeHtml(t.status)} <span>Fecha prevista: ${t.due?fd(t.due):'Sin fecha'}</span></p></article>`).join(''):'<p>No hay tareas abiertas vinculadas a estas actuaciones.</p>'}</div></div>`;
}
const latamDependencyLoads=new Map();
function loadLatamReportDependency(url){
  if(!latamDependencyLoads.has(url))latamDependencyLoads.set(url,new Promise((resolve,reject)=>{const script=document.createElement('script');script.src=url;script.onload=resolve;script.onerror=()=>{latamDependencyLoads.delete(url);script.remove();reject(Error('No se pudo cargar el generador de PowerPoint. Vuelve a intentarlo.'));};document.head.appendChild(script);}));
  return latamDependencyLoads.get(url);
}
async function exportLatamPresentation(){
  if(latamReportUI.busy)return;
  let report;try{report=currentLatamReport();if(!report.count)throw Error('No hay colaboraciones en el periodo seleccionado.');}catch(e){latamReportUI.error=e.message;render();return;}
  latamReportUI.busy=true;latamReportUI.error='';render();
  try{
    if(!window.JSZip)await loadLatamReportDependency('features/reforma/vendor/jszip-3.10.1.min.js');
    if(!window.PptxGenJS)await loadLatamReportDependency('features/reforma/vendor/pptxgenjs-3.12.0.min.js');
    await document.fonts.load('18px Inter');await document.fonts.load('600 19px Inter');
    const response=await fetch('features/reforma/assets/cegid.png');if(!response.ok)throw Error('No se pudo cargar el logo Cegid.');
    const logo=await new Promise((resolve,reject)=>{const reader=new FileReader();reader.onload=()=>resolve(reader.result);reader.onerror=()=>reject(Error('No se pudo preparar el logo.'));response.blob().then(b=>reader.readAsDataURL(b),reject);});
    const blob=await BDLatamReport.presentation(report,window.PptxGenJS,logo,window.JSZip),url=URL.createObjectURL(blob),link=document.createElement('a');link.href=url;link.download=report.filename;link.click();setTimeout(()=>URL.revokeObjectURL(url),30000);
    showToast('PowerPoint generado','Informe ejecutivo del periodo seleccionado.');
  }catch(e){latamReportUI.error=e.message||'No se pudo generar el informe.';}
  finally{latamReportUI.busy=false;render();}
}
