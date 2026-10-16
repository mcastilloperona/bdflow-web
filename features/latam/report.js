(function(root,factory){if(typeof module==='object'&&module.exports)module.exports=factory();else root.BDLatamReport=factory();})(typeof globalThis!=='undefined'?globalThis:this,function(){
  'use strict';
  const clean=v=>String(v??'').trim();
  const hours=v=>Number.isFinite(Number(v))?Number(v):0;
  const round=v=>Math.round((v+Number.EPSILON)*100)/100;
  const date=v=>/^\d{4}-\d{2}-\d{2}$/.test(clean(v))?clean(v):'';
  const fmtDate=v=>date(v)?v.split('-').reverse().join('/'):'Sin fecha';
  const number=v=>new Intl.NumberFormat('es-ES',{maximumFractionDigits:2}).format(v);
  function model(actions,tasks,filters={}){
    const from=date(filters.from),to=date(filters.to);
    if(from&&to&&from>to)throw Error('La fecha inicial debe ser anterior o igual a la fecha final.');
    const rows=(actions||[]).filter(a=>(!from||(date(a.date)&&a.date>=from))&&(!to||(date(a.date)&&a.date<=to))).map(a=>({...a,task:clean(a.task)||'Sin descripción',collaboration:clean(a.collaboration)||'Sin especificar',context:clean(a.context),hours:hours(a.hours)})).sort((a,b)=>clean(a.date+' '+a.time).localeCompare(clean(b.date+' '+b.time))||clean(a.id).localeCompare(clean(b.id)));
    const groups=new Map();for(const row of rows){const group=groups.get(row.collaboration)||{name:row.collaboration,count:0,hours:0};group.count++;group.hours=round(group.hours+row.hours);groups.set(row.collaboration,group);}
    const taskMap=new Map((tasks||[]).map(t=>[t.id,t])),linkedIds=new Set(rows.map(a=>a.linkedTaskId).filter(Boolean));
    const linked=[...linkedIds].map(id=>taskMap.get(id)).filter(Boolean);
    const next=linked.filter(t=>t.status!=='Finalizada').map(t=>({id:t.id,title:clean(t.title)||'Sin título',next:clean(t.next),due:date(t.due),status:clean(t.status)||'Sin estado'})).sort((a,b)=>(a.due||'9999').localeCompare(b.due||'9999')||a.title.localeCompare(b.title));
    const first=rows.map(a=>date(a.date)).filter(Boolean).sort()[0]||'',last=rows.map(a=>date(a.date)).filter(Boolean).sort().at(-1)||'';
    const start=from||first,end=to||last;
    return {rows,groups:[...groups.values()].sort((a,b)=>b.count-a.count||b.hours-a.hours||a.name.localeCompare(b.name)),count:rows.length,totalHours:round(rows.reduce((s,a)=>s+a.hours,0)),withContext:rows.filter(a=>a.context).length,linkedCount:linked.length,next,from,to,start,end,period:start&&end?(start===end?fmtDate(start):fmtDate(start)+' al '+fmtDate(end)):'Todo el histórico',filename:'Seguimiento LATAM'+(start&&end?' '+start.split('-').reverse().join('-')+(end!==start?' al '+end.split('-').reverse().join('-'):''):'')+'.pptx'};
  }
  // Wrap real content into complete lines, then paginate without discarding text.
  function wrap(value,width=850,size=18,measure=t=>String(t).length*size*.62){
    const lines=[];for(const paragraph of String(value??'').replace(/\r/g,'').split('\n')){
      if(!paragraph){lines.push('');continue;}
      let line='';for(const word of paragraph.split(/\s+/)){
        if(measure(line?line+' '+word:word)<=width){line=line?line+' '+word:word;continue;}
        if(line){lines.push(line);line='';}
        let part='';for(const ch of word){if(part&&measure(part+ch)>width){lines.push(part);part='';}part+=ch;}line=part;
      }if(line)lines.push(line);
    }return lines;
  }
  function pages(blocks,measure){
    const result=[];let page=[];
    function flush(){if(page.length){result.push(page);page=[];}}
    for(const block of blocks){
      const lines=block.lines.flatMap(l=>wrap(l.text,850,l.size||18,t=>measure(t,l.size||18,l.bold)).map(text=>({...l,text})));
      if(page.length&&lines.length<=16&&page.length+lines.length+1>16)flush();
      if(page.length>=16)flush();
      if(page.length)page.push({text:'',size:18});
      for(const line of lines){if(page.length>=16){flush();page.push({text:block.label+' (continuación)',bold:true,size:18});}page.push(line);}
    }flush();return result;
  }
  function slidePlan(report,measure=(text,size)=>String(text).length*size*.62){
    if(!report.count)throw Error('No hay colaboraciones en el periodo seleccionado.');
    const slides=[{type:'summary',title:'Colaboración LATAM'}];
    const groupBlocks=report.groups.map(g=>({label:g.name,lines:[{text:g.name,bold:true},{text:'Actuaciones: '+g.count+'   Apoyo: '+number(g.hours)+' h'}]}));
    for(const lines of pages(groupBlocks,measure))slides.push({type:'body',title:'Colaboración por persona o equipo',lines});
    const detailBlocks=report.rows.map((a,i)=>({label:'Actuación '+(i+1),lines:[{text:fmtDate(a.date)+(a.time?'   '+a.time:'')+'   '+a.collaboration+'   Apoyo: '+number(a.hours)+' h',bold:true,size:18},{text:a.task,bold:true,size:19},{text:'Comentario o resultado: '+(a.context||'Sin registrar'),size:18}]}));
    for(const lines of pages(detailBlocks,measure))slides.push({type:'body',title:'Actuaciones y resultados registrados',lines});
    if(report.next.length){
      const nextBlocks=report.next.map(t=>({label:t.title,lines:[{text:t.title,bold:true,size:19},{text:'Próximo paso: '+(t.next||'Sin registrar'),size:18},{text:'Estado actual: '+t.status+'   Fecha prevista: '+(t.due?fmtDate(t.due):'Sin fecha'),size:18}]}));
      for(const lines of pages(nextBlocks,measure))slides.push({type:'body',title:'Próximos pasos de tareas vinculadas',lines,footnote:'Estado actual de las tareas vinculadas a la actividad del periodo.'});
    }else slides.push({type:'body',title:'Próximos pasos',lines:[{text:'No hay tareas abiertas vinculadas a las actuaciones de este periodo.',size:22},{text:'Los registros conservan los comentarios o resultados que hayas añadido.',size:18}]});
    return slides;
  }
  async function presentation(report,PptxGenJS,logo,JSZip){
    const ctx=document.createElement('canvas').getContext('2d');
    const measure=(text,size,bold)=>{ctx.font=(bold?'600 ':'400 ')+size+'px Inter, Arial';return ctx.measureText(text).width;};
    const plans=slidePlan(report,measure),pptx=new PptxGenJS();pptx.layout='LAYOUT_WIDE';pptx.author='BDFlow';pptx.company='Cegid';pptx.subject='Actividad y colaboración LATAM';pptx.title='Seguimiento LATAM '+report.period;pptx.lang='es-ES';pptx.theme={headFontFace:'Inter',bodyFontFace:'Inter',lang:'es-ES'};
    const colors={blue:'2962FF',navy:'123E69',gray:'64748B'};
    function text(slide,value,x,y,w,h,size=18,color=colors.navy,bold=false){slide.addText(String(value),{x,y,w,h,fontFace:'Inter',fontSize:size,color,bold,margin:0,breakLine:false,vertAnchor:'top',paraSpaceAfterPt:0});}
    plans.forEach((plan,index)=>{const slide=pptx.addSlide();slide.background={color:'FFFFFF'};if(logo)slide.addImage({data:logo,x:.5,y:.26,w:1.1,h:.45});text(slide,plan.title,1.93,.34,10.9,.6,30,colors.navy,true);text(slide,report.period,.5,1.07,12.3,.3,14,colors.gray);
      if(plan.type==='summary'){
        text(slide,'Actividad y seguimiento para Dirección',.5,1.68,12.2,.52,27,colors.navy,true);
        [[String(report.count),'Actuaciones registradas'],[String(report.groups.length),'Personas, equipos o áreas'],[number(report.totalHours)+' h','Tiempo de apoyo']].forEach(([value,label],i)=>{const x=.5+i*4.2;text(slide,value,x,2.66,4,1,49,colors.blue,true);text(slide,label,x,3.72,4,.42,18,colors.gray);});
        text(slide,'Actuaciones con comentario o resultado: '+report.withContext+' de '+report.count+'.',.5,4.68,12, .55,22,colors.navy,true);
        text(slide,'Tareas BDFlow vinculadas: '+report.linkedCount+'. Abiertas: '+report.next.length+'.',.5,5.4,12,.5,20,colors.navy);
        text(slide,'El detalle incluye la actividad del periodo y los próximos pasos de las tareas vinculadas.',.5,6.2,12,.6,18,colors.gray);
      }else plan.lines.forEach((line,i)=>text(slide,line.text,.5,1.65+i*.30,12,.30,line.size||18,colors.navy,!!line.bold));
      if(plan.footnote)text(slide,plan.footnote,.5,6.77,11.8,.28,12,colors.gray);
      text(slide,'Seguimiento LATAM',.5,7.1,10,.2,10,colors.gray);text(slide,(index+1)+' / '+plans.length,12,7.1,.8,.2,10,colors.gray);
    });
    const blob=await pptx.write({outputType:'blob',compression:true});
    // This browser build can declare unused masters. Keep only package parts that exist.
    const zip=await JSZip.loadAsync(blob),types=await zip.file('[Content_Types].xml').async('string');
    zip.file('[Content_Types].xml',types.replace(/<Override\b[^>]*\/>/g,tag=>{const part=tag.match(/\bPartName="([^"]+)"/)?.[1];return part&&!zip.file(part.replace(/^\//,''))?'':tag;}));
    return zip.generateAsync({type:'blob',compression:'DEFLATE',mimeType:'application/vnd.openxmlformats-officedocument.presentationml.presentation'});
  }
  return {model,wrap,pages,slidePlan,presentation,fmtDate,number};
});
