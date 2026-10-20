window.BDExecutiveTemplate=(function(){
 'use strict';
 const NS=914400,wide=13.3333333333,high=7.5;
 const enc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]));
 const dec=v=>String(v??'').replace(/&#x([0-9a-f]+);/gi,(_,n)=>String.fromCodePoint(parseInt(n,16))).replace(/&#(\d+);/g,(_,n)=>String.fromCodePoint(+n)).replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"').replace(/&apos;/g,"'").replace(/&amp;/g,'&');
 function attr(xml,name){return dec(xml.match(new RegExp('(?:^|\\s)'+name+'="([^"]*)"'))?.[1]||'');}
 function bounds(shape){const x=shape.match(/<a:xfrm\b[^>]*>([\s\S]*?)<\/a:xfrm>/)?.[1]||'',off=x.match(/<a:off\b[^>]*\/>/)?.[0]||'',ext=x.match(/<a:ext\b[^>]*\/>/)?.[0]||'';return {x:+attr(off,'x')/NS,y:+attr(off,'y')/NS,w:+attr(ext,'cx')/NS,h:+attr(ext,'cy')/NS};}
 function getText(shape){return [...shape.matchAll(/<a:p\b[^>]*>([\s\S]*?)<\/a:p>/g)].map(p=>[...p[1].matchAll(/<a:t\b[^>]*>([\s\S]*?)<\/a:t>/g)].map(t=>dec(t[1])).join('')).join('\n');}
 function setText(shape,value){return shape.replace(/(<p:txBody\b[^>]*>)([\s\S]*?)(<\/p:txBody>)/,(_,open,body,close)=>{const bodyPr=body.match(/<a:bodyPr\b[^>]*(?:\/>|>[\s\S]*?<\/a:bodyPr>)/)?.[0]||'<a:bodyPr/>',list=body.match(/<a:lstStyle\b[^>]*(?:\/>|>[\s\S]*?<\/a:lstStyle>)/)?.[0]||'<a:lstStyle/>',first=body.match(/<a:p\b[^>]*>([\s\S]*?)<\/a:p>/)?.[1]||'',pPr=first.match(/<a:pPr\b[^>]*(?:\/>|>[\s\S]*?<\/a:pPr>)/)?.[0]||'',rPr=first.match(/<a:rPr\b[^>]*(?:\/>|>[\s\S]*?<\/a:rPr>)/)?.[0]||first.match(/<a:endParaRPr\b[^>]*(?:\/>|>[\s\S]*?<\/a:endParaRPr>)/)?.[0]?.replaceAll('endParaRPr','rPr')||'';return open+bodyPr+list+String(value).split('\n').map(line=>'<a:p>'+pPr+'<a:r>'+rPr+'<a:t xml:space="preserve">'+enc(line)+'</a:t></a:r></a:p>').join('')+close;});}
 function place(shape,x,w){return shape.replace(/(<a:off\b[^>]*\bx=")\d+(")/,'$1'+Math.round(x*NS)+'$2').replace(/(<a:ext\b[^>]*\bcx=")\d+(")/,'$1'+Math.round(w*NS)+'$2');}
 function resolve(base,target){if(!target||/^[a-z]+:/i.test(target))throw Error('La plantilla contiene una referencia externa no admitida.');const out=target.startsWith('/')?[]:base.split('/').slice(0,-1);for(const item of target.split('/')){if(!item||item==='.')continue;if(item==='..'){if(!out.length)throw Error('Referencia de plantilla no válida.');out.pop();}else out.push(item);}return out.join('/');}
 function relsPath(path){const parts=path.split('/'),name=parts.pop();return parts.join('/')+'/_rels/'+name+'.rels';}
 function relationship(xml,id){return [...xml.matchAll(/<Relationship\b[^>]*\/>/g)].map(v=>v[0]).find(t=>attr(t,'Id')===id)||null;}
 function slides(zip){return Object.keys(zip.files).filter(p=>/^ppt\/slides\/slide\d+\.xml$/.test(p)).sort((a,b)=>+a.match(/slide(\d+)/)[1]-+b.match(/slide(\d+)/)[1]);}
 async function orderedSlides(zip){const p='ppt/presentation.xml',xml=await zip.file(p).async('string'),rels=await zip.file(relsPath(p)).async('string');return [...xml.matchAll(/<p:sldId\b[^>]*\br:id="([^"]+)"[^>]*\/>/g)].map(v=>{const r=relationship(rels,v[1]);if(!r)throw Error('Orden de diapositivas no válido.');return resolve(p,attr(r,'Target'));});}
 async function validate(bytes,JSZip){if(bytes.byteLength>20*1024*1024)throw Error('La plantilla supera los 20 MB.');const zip=await JSZip.loadAsync(bytes);if(Object.keys(zip.files).length>3000||!zip.file('ppt/presentation.xml'))throw Error('Selecciona un PowerPoint PPTX válido.');const paths=await orderedSlides(zip);if(![3,4].includes(paths.length)||paths.length!==slides(zip).length)throw Error('Usa el PowerPoint ejecutivo de tres diapositivas, con portada opcional.');for(const path of paths){if(!zip.file(path)||(await zip.file(path).async('string')).length>5000000)throw Error('La plantilla es demasiado compleja.');}return zip;}
 function updateSlide(xml,index,m,scaleX=1,scaleY=1){
  const E=window.BDExecutive,signed=n=>n==null?'No disponible':(n>0?'+':'')+E.money(n),d=E.date,roles=index===0?[
   [.65,1.61,'Reforma Laboral México | '+d(m.first.date)+' a '+d(m.last.date)],
   [.88,2.78,'Inicio '+d(m.first.date)+(m.first.reconstructed?' (reconstruido)':'')],[5.03,2.78,'Actual '+d(m.last.date)],
   [.88,3.16,E.money(m.first.amount)],[5.03,3.16,E.money(m.last.amount)],
   [.88,3.77,m.first.count+' oportunidades'],[5.03,3.77,m.last.count+' oportunidades'],
   [.7,4.9,signed(m.delta)+' | '+E.pct(m.percent)],[.7,5.52,(m.countDelta>0?'+':'')+m.countDelta+' oportunidades desde el inicio'],
   [9.4,2.82,E.money(m.last.won)],[9.4,4.02,E.money(m.target)],[9.4,4.86,'Cumplimiento: '+E.pct(m.achievement)],[9.4,5.5,'Inicio: '+E.money(m.first.won)]
  ]:index===1?[
   [9.4,3,signed(m.delta)],[9.4,4.18,m.rises.length+' intervalos al alza\n'+m.declines.length+' intervalos a la baja'],
   [9.4,5.04,m.largestDecline?'Mayor bajada observada:\n'+signed(m.largestDecline.delta)+'\n'+d(m.largestDecline.from)+' a '+d(m.largestDecline.to):'Sin bajadas observadas durante el periodo']
  ]:[
   [9.4,3.12,E.money(m.forecastTotal)],[9.4,3.92,m.concentration?E.pct(m.concentrationPercent)+' concentrado en '+new Intl.DateTimeFormat('es-ES',{month:'long',timeZone:'UTC'}).format(new Date(m.concentration.month+'-01T00:00:00Z')):'Sin concentración de importe disponible']
  ];
  const matches=[...xml.matchAll(/<p:sp\b[\s\S]*?<\/p:sp>/g)],used=new Set(),replacements=new Map();
  const names=index===0?[2,7,11,8,12,9,13,15,16,19,21,22,23]:index===1?[7,9,10]:[10,11];
  for(const [role,[x,y,value]] of roles.entries()){let best=matches.findIndex((match,i)=>!used.has(i)&&attr(match[0].match(/<p:cNvPr\b[^>]*\/>/)?.[0]||'','name')==='Text '+names[role]);let score=Infinity;if(best<0)matches.forEach((match,i)=>{if(used.has(i)||!getText(match[0]))return;const b=bounds(match[0]),dx=Math.abs(b.x/scaleX-x),dy=Math.abs(b.y/scaleY-y);if(dx<.32&&dy<.15&&dx+dy<score){best=i;score=dx+dy;}});if(best>=0){used.add(best);replacements.set(best,setText(matches[best][0],value));}}
  if(index===1){matches.forEach((match,i)=>{const b=bounds(match[0]),t=getText(match[0]);if(b.y/scaleY<2&&/fechas guardadas/.test(t))replacements.set(i,setText(match[0],t.replace(/\d+ fechas guardadas/,m.points.length+' fechas guardadas')));});}
  if(index===2){const counters=matches.map((v,i)=>({xml:v[0],i,b:bounds(v[0]),text:getText(v[0])})).filter(v=>/^\d+\s+(oportunidades|opps)$/.test(v.text.trim())&&v.b.x/scaleX<9&&Math.abs(v.b.y/scaleY-5.6)<.25).sort((a,b)=>a.b.x-b.b.x);const template=counters[0];if(template){const maxId=Math.max(100,...[...xml.matchAll(/<p:cNvPr\b[^>]*\bid="(\d+)"/g)].map(v=>+v[1]));counters.forEach((v,i)=>{if(i<m.forecast.length){let s=place(setText(v.xml,m.forecast[i].count+(m.forecast.length>4?' opps':' oportunidades')),(.8+i*7.8/m.forecast.length)*scaleX,7.8/m.forecast.length*scaleX);replacements.set(v.i,s);}else replacements.set(v.i,'');});let extra='';for(let i=counters.length;i<m.forecast.length;i++){let s=place(setText(template.xml,m.forecast[i].count+(m.forecast.length>4?' opps':' oportunidades')),(.8+i*7.8/m.forecast.length)*scaleX,7.8/m.forecast.length*scaleX).replace(/(<p:cNvPr\b[^>]*\bid=")\d+(")/,'$1'+(maxId+i+1)+'$2');extra+=s;}if(extra)xml=xml.replace('</p:spTree>',extra+'</p:spTree>');}}
  for(const [k,match] of matches.entries()){if(/\d[\d., \u00A0]*\s*(?:€|%)/.test(getText(match[0]))&&!replacements.has(k))throw Error('La plantilla contiene un importe que no se puede actualizar automáticamente. Conserva la posición de los datos de la presentación ejecutiva.');}let i=0;return xml.replace(/<p:sp\b[\s\S]*?<\/p:sp>/g,shape=>{const n=i++;return replacements.has(n)?replacements.get(n):shape;});
 }
 function updateChart(xml,labels,values){
  const cache=(tag,items,numeric)=>'<c:'+tag+'>'+(numeric?(xml.match(/<c:formatCode>[\s\S]*?<\/c:formatCode>/)?.[0]||'<c:formatCode>General</c:formatCode>'):'')+'<c:ptCount val="'+items.length+'"/>'+items.map((v,i)=>v==null?'':'<c:pt idx="'+i+'"><c:v>'+enc(v)+'</c:v></c:pt>').join('')+'</c:'+tag+'>';
  xml=xml.replace(/<c:cat\b[^>]*>[\s\S]*?<\/c:cat>/,block=>block.replace(/<c:strCache\b[^>]*>[\s\S]*?<\/c:strCache>/,cache('strCache',labels,false)).replace(/(<c:f>[\s\S]*?\$[A-Z]+\$\d+:\$[A-Z]+\$)\d+(<\/c:f>)/,'$1'+Math.max(2,labels.length+1)+'$2'));
  xml=xml.replace(/<c:val\b[^>]*>[\s\S]*?<\/c:val>/,block=>block.replace(/<c:numCache\b[^>]*>[\s\S]*?<\/c:numCache>/,cache('numCache',values,true)).replace(/(<c:f>[\s\S]*?\$[A-Z]+\$\d+:\$[A-Z]+\$)\d+(<\/c:f>)/,'$1'+Math.max(2,values.length+1)+'$2'));
  const skip=Math.max(1,Math.ceil(labels.length/7));xml=xml.replace(/<c:catAx\b[^>]*>[\s\S]*?<\/c:catAx>/,block=>/<c:tickLblSkip\b/.test(block)?block.replace(/<c:tickLblSkip\b[^>]*\/>/,'<c:tickLblSkip val="'+skip+'"/>'):block.replace('</c:catAx>','<c:tickLblSkip val="'+skip+'"/></c:catAx>'));return xml;
 }
 async function chart(zip,path,labels,values,JSZip,ExcelJS){
  const file=zip.file(path);if(!file)throw Error('Falta un gráfico en la plantilla.');let xml=await file.async('string');const updated=updateChart(xml,labels,values);if(!updated.includes('<c:ptCount val="'+labels.length+'"/>'))throw Error('El gráfico de la plantilla no es compatible.');zip.file(path,updated);
  const relFile=zip.file(relsPath(path));if(!relFile)return;const relXML=await relFile.async('string'),id=xml.match(/<c:externalData\b[^>]*\br:id="([^"]+)"/)?.[1];if(!id)return;const rel=relationship(relXML,id);if(!rel)return;const workbookPath=resolve(path,attr(rel,'Target'));if(!zip.file(workbookPath))return;const book=new ExcelJS.Workbook(),sheetName=dec(xml.match(/<c:f>([^!]+)!\$[A-Z]/)?.[1]||'Sheet1').replace(/^'|'$/g,'').replace(/''/g,"'"),sheet=book.addWorksheet(sheetName);sheet.addRow(['Fecha','SaaS Booking abierto sin G']);labels.forEach((label,i)=>sheet.addRow([label,values[i]??null]));zip.file(workbookPath,await book.xlsx.writeBuffer());
 }
 async function imagePagination(zip,path,xml){
  const rel=zip.file(relsPath(path));if(!rel)return null;const relXML=await rel.async('string');
  for(const pic of xml.matchAll(/<p:pic\b[\s\S]*?<\/p:pic>/g)){const b=bounds(pic[0]);if(b.w<wide*.9||b.h<high*.9)continue;const id=pic[0].match(/\br:embed="([^"]+)"/)?.[1],r=id?relationship(relXML,id):null;if(!r)continue;const media=resolve(path,attr(r,'Target')),file=zip.file(media);if(!file)continue;const mime=/\.jpe?g$/i.test(media)?'image/jpeg':'image/png';return window.BDExecutive.templatePagination('data:'+mime+';base64,'+await file.async('base64'));}
  return null;
 }
 function numberSlide(xml,index,pagination,scaleX=1,scaleY=1){
  let existing=false;xml=xml.replace(/<p:sp\b[\s\S]*?<\/p:sp>/g,shape=>{const nv=shape.match(/<p:cNvPr\b[^>]*\/>/)?.[0]||'';if(attr(nv,'name')==='Número de diapositiva'){existing=true;return setText(shape,index);}return shape;});if(existing)return xml;
  xml=xml.replace(/<p:sp\b[\s\S]*?<\/p:sp>/g,shape=>{const b=bounds(shape),t=getText(shape).trim();return b.y/scaleY>6.4&&/^\d+(?:\s*\/\s*\d+)?$/.test(t)?'':shape;});
  const maxId=Math.max(100,...[...xml.matchAll(/<p:cNvPr\b[^>]*\bid="(\d+)"/g)].map(v=>+v[1])),p=pagination?.patch,n=pagination?.number||{x:12.15,y:6.55,w:.5,h:.24,fontSize:9,color:'64748B'};
  const transform=(b)=>'<a:xfrm><a:off x="'+Math.round(b.x*scaleX*NS)+'" y="'+Math.round(b.y*scaleY*NS)+'"/><a:ext cx="'+Math.round(b.w*scaleX*NS)+'" cy="'+Math.round(b.h*scaleY*NS)+'"/></a:xfrm>';
  let shapes=p?'<p:sp><p:nvSpPr><p:cNvPr id="'+(maxId+1)+'" name="Fondo numeración dinámica"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>'+transform(p)+'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="'+p.color+'"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr></p:sp>':'';
  shapes+='<p:sp><p:nvSpPr><p:cNvPr id="'+(maxId+2)+'" name="Número de diapositiva"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>'+transform(n)+'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="es-ES" sz="'+Math.round(n.fontSize*100)+'"><a:solidFill><a:srgbClr val="'+n.color+'"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>'+index+'</a:t></a:r></a:p></p:txBody></p:sp>';return xml.replace('</p:spTree>',shapes+'</p:spTree>');
 }
 async function presentation(m,bytes,JSZip,ExcelJS,options={}){
  const zip=await validate(bytes,JSZip),pres=await zip.file('ppt/presentation.xml').async('string'),size=pres.match(/<p:sldSz\b[^>]*\/>/)?.[0]||'',scaleX=(+attr(size,'cx')/NS||wide)/wide,scaleY=(+attr(size,'cy')/NS||high)/high;
  const paths=await orderedSlides(zip),offset=paths.length===4?1:0;
  for(const [position,path] of paths.entries()){
   let xml=await zip.file(path).async('string');if(position<offset){const month=new Intl.DateTimeFormat('es-ES',{month:'short',timeZone:'UTC'}).format(new Date(m.last.date+'T00:00:00Z')).replace('.','').toUpperCase(),label='RLM '+month+' '+m.last.date.slice(2,4);xml=xml.replace(/<a:t\b[^>]*>[\s\S]*?<\/a:t>/g,t=>t.replace(/RLM\s+[A-ZÁÉÍÓÚ]{3,4}\s+\d{2}/,label));zip.file(path,xml);continue;}const i=position-offset;const pagination=/name="Número de diapositiva"/.test(xml)?null:await imagePagination(zip,path,xml);xml=updateSlide(xml,i,m,scaleX,scaleY);xml=numberSlide(xml,i+1,pagination,scaleX,scaleY);zip.file(path,xml);
   if(i===0)continue;const relFile=zip.file(relsPath(path));if(!relFile)throw Error('Falta un gráfico en la plantilla.');const relXML=await relFile.async('string'),refs=[...xml.matchAll(/<c:chart\b[^>]*\br:id="([^"]+)"/g)];if(!refs.length)throw Error('La plantilla necesita gráficos editables en las diapositivas 2 y 3.');
   const labels=i===1?m.points.map(v=>window.BDExecutive.date(v.date)):m.forecast.map(v=>new Intl.DateTimeFormat('es-ES',{month:'long',year:'numeric',timeZone:'UTC'}).format(new Date(v.month+'-01T00:00:00Z'))),values=i===1?m.points.map(v=>v.amount):m.forecast.map(v=>v.amount);
   for(const ref of refs){const r=relationship(relXML,ref[1]);if(!r)throw Error('Referencia de gráfico no válida.');await chart(zip,resolve(path,attr(r,'Target')),labels,values,JSZip,ExcelJS);}
  }
  return zip.generateAsync({type:options.outputType||'blob',compression:'DEFLATE'});
 }
 return {validate,presentation,updateSlide,updateChart,getText,setText,bounds,numberSlide,resolve,orderedSlides};
})();
