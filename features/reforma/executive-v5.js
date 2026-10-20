window.BDExecutive=(function(){
 'use strict';
 const money=n=>n==null?'No disponible':new Intl.NumberFormat('es-ES',{style:'currency',currency:'EUR',maximumFractionDigits:2}).format(n);
 const date=d=>d.split('-').reverse().join('/');
 const pct=n=>n==null?'No disponible':n.toLocaleString('es-ES',{maximumFractionDigits:1})+' %';
 function model(series,months,target=1500000){
  const points=series.slice().sort((a,b)=>a.date.localeCompare(b.date));if(!points.length)throw Error('No hay listados disponibles.');
  const first=points[0],last=points.at(-1),known=first.amount!=null&&last.amount!=null,delta=known?Math.round((last.amount-first.amount)*100)/100:null;
  const moves=points.slice(1).map((p,i)=>({from:points[i].date,to:p.date,delta:p.amount==null||points[i].amount==null?null:Math.round((p.amount-points[i].amount)*100)/100}));
  const declines=moves.filter(v=>v.delta<0),rises=moves.filter(v=>v.delta>0),largestDecline=declines.slice().sort((a,b)=>a.delta-b.delta)[0]||null;
  const forecast=months.slice().sort((a,b)=>a.month.localeCompare(b.month)),forecastKnown=forecast.every(v=>v.amount!=null),forecastTotal=forecastKnown?Math.round(forecast.reduce((n,v)=>n+v.amount,0)*100)/100:null;
  if(last.amount!=null&&forecastTotal!=null&&Math.abs(last.amount-forecastTotal)>0.01)throw Error('Los cierres previstos no cuadran con el pipeline seleccionado.');
  const concentration=forecastKnown&&forecastTotal>0?forecast.slice().sort((a,b)=>b.amount-a.amount)[0]:null;
  return {points,first,last,delta,percent:known&&first.amount!==0?delta/first.amount*100:null,countDelta:last.count-first.count,target,achievement:last.won==null?null:last.won/target*100,moves,declines,rises,largestDecline,forecast,forecastTotal,concentration,concentrationPercent:concentration?concentration.amount/forecastTotal*100:null,reconstructed:points.some(p=>p.reconstructed)};
 }

 function findTemplateNumber(data,width,height,fullWidth,fullHeight,startY){
  const distance=(i,j)=>Math.abs(data[i]-data[j])+Math.abs(data[i+1]-data[j+1])+Math.abs(data[i+2]-data[j+2]);
  const mask=new Uint8Array(width*height),seen=new Uint8Array(width*height),components=[];
  for(let y=0;y<height;y++)for(let x=1;x<width;x++)if(distance((y*width+x)*4,y*width*4)>110)mask[y*width+x]=1;
  for(let p=0;p<mask.length;p++){if(!mask[p]||seen[p])continue;const queue=[p];seen[p]=1;let minX=width,maxX=0,minY=height,maxY=0,dark=p,contrast=0;for(let q=0;q<queue.length;q++){const n=queue[q],x=n%width,y=Math.floor(n/width);minX=Math.min(minX,x);maxX=Math.max(maxX,x);minY=Math.min(minY,y);maxY=Math.max(maxY,y);const d=distance(n*4,y*width*4);if(d>contrast){contrast=d;dark=n;}for(let dy=-1;dy<=1;dy++)for(let dx=-1;dx<=1;dx++){const nx=x+dx,ny=y+dy;if(nx<0||nx>=width||ny<0||ny>=height)continue;const i=ny*width+nx;if(mask[i]&&!seen[i]){seen[i]=1;queue.push(i);}}}
   const w=maxX-minX+1,h=maxY-minY+1;if(minX>=fullWidth*.003&&minX<fullWidth*.05&&startY+minY>fullHeight*.88&&h>=Math.max(6,fullHeight*.007)&&h<fullHeight*.06&&w<=Math.max(15,fullWidth*.03)&&w/h<=.95&&queue.length>=5)components.push({minX,maxX,minY,maxY,w,h,dark});
  }
  const hex=i=>[data[i],data[i+1],data[i+2]].map(v=>v.toString(16).padStart(2,'0')).join('').toUpperCase();
  const c=components.sort((a,b)=>a.minX-b.minX||b.h-a.h)[0];
  if(!c){const color=hex(Math.min(height-1,Math.floor(height*.65))*width*4);return {patch:{x:.08,y:7.03,w:.52,h:.37,color},number:{x:.14,y:7.06,w:.35,h:.27,fontSize:10,color:'123E69'}};}
  const x=Math.max(0,c.minX-3),y=Math.max(0,c.minY-3),w=c.w+6,h=c.h+6;
  return {patch:{x:x/fullWidth*13.3333333333,y:(startY+y)/fullHeight*7.5,w:w/fullWidth*13.3333333333,h:h/fullHeight*7.5,color:hex(c.minY*width*4)},number:{x:(c.minX-4)/fullWidth*13.3333333333,y:(startY+c.minY-3)/fullHeight*7.5,w:(c.w+8)/fullWidth*13.3333333333,h:(c.h+6)/fullHeight*7.5,fontSize:Math.max(7,Math.min(18,c.h/fullHeight*7.5*72)),color:hex(c.dark*4)}};
 }
 async function templatePagination(background){
  const img=new Image();img.src=background;await img.decode();const canvas=document.createElement('canvas');canvas.width=img.naturalWidth;canvas.height=img.naturalHeight;const g=canvas.getContext('2d');if(!g)throw Error('No se pudo preparar la numeración de la plantilla.');g.fillStyle='#FFFFFF';g.fillRect(0,0,canvas.width,canvas.height);g.drawImage(img,0,0);const top=Math.floor(canvas.height*.86),w=Math.max(1,Math.floor(canvas.width*.08)),h=canvas.height-top;const pixels=g.getImageData(0,top,w,h).data;return findTemplateNumber(pixels,w,h,canvas.width,canvas.height,top);
 }

 async function presentation(m,PptxGenJS,background,logoData=null,pagination=null){
  const pptx=new PptxGenJS();pptx.layout='LAYOUT_WIDE';pptx.author='BDFlow';pptx.company='Cegid';pptx.lang='es-ES';pptx.title='Evolución ejecutiva Reforma Laboral '+date(m.last.date);
  const navy='123E69',blue='2962FF',gray='64748B',red='B91C1C',orange='FF5C35';
  const text=(s,t,x,y,w,h,size=15,opts={})=>s.addText(String(t),{x,y,w,h,fontFace:'Arial',fontSize:size,color:navy,margin:0,breakLine:false,valign:'mid',...opts});
  const box=(s,x,y,w,h,color='FFFFFF')=>s.addShape(pptx.ShapeType.rect,{x,y,w,h,line:{color},fill:{color}});
  const footer=m.reconstructed?'Base inicial identificada como reconstruida en BDFlow. Fechas sin listado no se rellenan.':'Solo listados guardados. Las fechas sin listado no se rellenan.';
  function slide(title,subtitle,index){const s=pptx.addSlide();s.background={color:'FFFFFF'};if(background)s.addImage({data:background,x:0,y:0,w:13.3333333333,h:7.5});else if(logoData)s.addImage({data:logoData,x:0.65,y:0.24,w:1.2,h:0.49});box(s,0.48,0.94,12.36,5.87);text(s,title,0.65,1.04,12,0.5,26,{bold:true,color:blue});text(s,subtitle,0.65,1.61,12,0.32,12,{color:gray});text(s,footer,0.65,6.55,11.4,0.24,9,{color:gray});if(background&&pagination){const p=pagination.patch,n=pagination.number;box(s,p.x,p.y,p.w,p.h,p.color);text(s,String(index),n.x,n.y,n.w,n.h,n.fontSize,{align:'center',color:n.color});}else{text(s,String(index),12.15,6.55,0.5,0.24,9,{align:'right',color:gray});}return s;}
  const signed=n=>n==null?'No disponible':(n>0?'+':'')+money(n);
  let s=slide('El plan: inicio y situación actual','Reforma Laboral México | '+date(m.first.date)+' a '+date(m.last.date),1);
  text(s,'SaaS Booking abierto sin etapa G',0.7,2.12,7.6,0.34,18,{bold:true});
  [['Inicio '+date(m.first.date),money(m.first.amount),m.first.count+' oportunidades'],['Actual '+date(m.last.date),money(m.last.amount),m.last.count+' oportunidades']].forEach((v,i)=>{const x=.7+i*4.15;box(s,x,2.64,3.9,1.5,'EEF3FA');text(s,v[0],x+.18,2.78,3.5,.25,13);text(s,v[1],x+.18,3.16,3.5,.42,25,{bold:true,color:blue});text(s,v[2],x+.18,3.77,3.5,.22,13);});
  text(s,'Variación desde el inicio',.7,4.42,7.8,.3,16,{bold:true});text(s,signed(m.delta)+'  |  '+pct(m.percent),.7,4.9,7.8,.47,25,{bold:true,color:blue});text(s,(m.countDelta>0?'+':'')+m.countDelta+' oportunidades desde el inicio',.7,5.52,7.8,.3,15);
  box(s,9.18,2.12,3.35,3.96,'EEF3FA');text(s,'Booking ganado 2026',9.4,2.32,2.92,.3,14);text(s,money(m.last.won),9.4,2.82,2.92,.4,24,{bold:true,color:blue});text(s,'Objetivo 2026',9.4,3.55,2.92,.3,14);text(s,money(m.target),9.4,4.02,2.92,.4,24,{bold:true});text(s,'Cumplimiento: '+pct(m.achievement),9.4,4.86,2.92,.35,16,{bold:true});text(s,'Inicio: '+money(m.first.won),9.4,5.5,2.92,.3,12,{color:gray});
  s=slide('La evolución durante todo el periodo','SaaS Booking abierto sin etapa G | '+m.points.length+' fechas guardadas',2);
  const plotted=m.points.filter(p=>Number.isFinite(p.amount));if(plotted.length)s.addChart(pptx.ChartType.line,[{name:'SaaS Booking abierto sin G',labels:plotted.map(p=>date(p.date)),values:plotted.map(p=>p.amount)}],{x:.65,y:2.1,w:8.3,h:3.95,catAxisLabelFontSize:10,catAxisLabelFreq:Math.max(1,Math.ceil(plotted.length/7)),valAxisLabelFontSize:10,showLegend:false,showTitle:false,showValue:false,chartColors:[blue],lineSize:3,showMarker:true,markerSize:5,valAxisMinVal:0,valAxisNumFmt:'#,##0 "€"',showCatName:false,showBorder:false,showCatName:false,showShadow:false,showCatName:false});
  box(s,9.18,2.14,3.35,3.95,'EEF3FA');text(s,'Lectura del periodo',9.4,2.34,2.9,.4,17,{bold:true});text(s,signed(m.delta),9.4,3.0,2.9,.4,24,{color:blue,bold:true});text(s,'Variación neta desde el inicio',9.4,3.49,2.9,.55,13);text(s,m.rises.length+' intervalos al alza\n'+m.declines.length+' intervalos a la baja',9.4,4.18,2.9,.7,14);if(m.largestDecline)text(s,'Mayor bajada observada:\n'+signed(m.largestDecline.delta)+'\n'+date(m.largestDecline.from)+' a '+date(m.largestDecline.to),9.4,5.04,2.9,.85,12,{color:red});
  s=slide('Pipeline por fecha de cierre prevista','Fechas declaradas en Salesforce | Sin ponderación por probabilidad',3);
  const forecastPlot=m.forecast.filter(v=>Number.isFinite(v.amount));if(forecastPlot.length)s.addChart(pptx.ChartType.bar,[{name:'SaaS Booking abierto sin G',labels:forecastPlot.map(v=>new Intl.DateTimeFormat('es-ES',{month:'long',year:'numeric',timeZone:'UTC'}).format(new Date(v.month+'-01T00:00:00Z'))),values:forecastPlot.map(v=>v.amount)}],{x:.65,y:2.1,w:8.2,h:3.35,barDir:'col',chartColors:[blue],showLegend:false,showValue:true,dataLabelPosition:'outEnd',dataLabelFormatCode:'#,##0 "€"',dataLabelFormatCodeSourceLinked:false,dataLabelFormatCode:'#,##0 "€"',dataLabelColor:navy,dataLabelFormatCodeSourceLinked:false,catAxisLabelFontSize:12,valAxisLabelFontSize:10,valAxisMinVal:0,valAxisNumFmt:'#,##0 "€"',showBorder:false,showShadow:false});
  if(m.forecast.length<=4)m.forecast.forEach((v,i)=>text(s,v.count+' oportunidades',.8+i*(7.8/Math.max(1,m.forecast.length)),5.6,7.8/Math.max(1,m.forecast.length),.3,13,{align:'center'}));else text(s,m.last.count+' oportunidades en '+m.forecast.length+' meses',.8,5.6,7.8,.3,13);
  box(s,9.18,2.14,3.35,3.95,'EEF3FA');text(s,'Pipeline abierto sin G',9.4,2.36,2.9,.55,15);text(s,money(m.forecastTotal),9.4,3.12,2.9,.42,24,{color:blue,bold:true});if(m.concentration){const month=new Intl.DateTimeFormat('es-ES',{month:'long',timeZone:'UTC'}).format(new Date(m.concentration.month+'-01T00:00:00Z'));text(s,pct(m.concentrationPercent)+' concentrado en '+month,9.4,3.92,2.9,.9,17,{bold:true,color:orange});}text(s,'Importes abiertos según su cierre previsto. No equivalen a Booking ganado.',9.4,5.12,2.9,.7,12,{color:gray});
  return pptx.write({outputType:'blob',compression:true});
 }
 return {model,presentation,templatePagination,findTemplateNumber,money,date,pct};
})();
