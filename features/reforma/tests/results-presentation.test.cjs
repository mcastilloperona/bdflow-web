const {test}=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),C=require('../core.js'),R=require('../results.js'),P=require('../prospecting.js'),D=require('../results-presentation.js'),Zip=require('../vendor/jszip-3.10.1.min.js'),Excel=require('../vendor/exceljs-4.4.0.min.js');
const imports=[{id:'s',report_date:'2026-09-29'},{id:'o',report_date:'2026-10-31'},{id:'n',report_date:'2026-11-17'},{id:'d',report_date:'2026-12-05'}];
const row=(id,imp,stage,close,amount,name='Newcli - Test')=>({id:id+imp,operation_id:id,import_id:imp,data:{account:'<Cuenta & prueba>',owner:'Propietario <Prueba>',opportunity:name,stage,close_date:close,target_plan:C.TARGET,amount,amount_basis:'saas'}});
const snapshots=[row('a','s','G - Sales Ready','2026-10-20',0),row('a','o','Closed Won','2026-10-20',100),row('a','n','Closed Won','2026-10-20',150),row('b','n','B - Negotiate','2026-11-30',250,'BI - Test'),row('a','d','Closed Won','2026-10-20',200),row('b','d','Closed Won','2026-12-04',350,'BI - Test')];
const calls=[{id:'p',report_date:'2026-11-17',entries_text:'2026-09-28 1\n2026-10-01 20\n2026-11-15 12',source_label:'Origen <test>',record_complete:false,notes:'Sin objetivo confirmado'},{id:'q',report_date:'2026-12-05',entries_text:'2026-12-04 35',source_label:'Otro corte',record_complete:true}];
const template=()=>fs.readFileSync(require.resolve('../assets/results-template.pptx'));
test('ranking uses actual named calls and opportunity cohorts without attributing calls to sales',()=>{
 const m=R.model(imports,snapshots,'n');m.ownerConversion=[{owner:'Ana María PEREZ',worked:10,converted:2,rate:0.2},{owner:'Luis DIAZ',worked:4,converted:0,rate:0}];
 m.ownerProgress=[{owner:'Ana María PEREZ',worked:10,advanced:5,steps:12},{owner:'Luis DIAZ',worked:4,advanced:1,steps:2}];
 const named=[{report_date:'2026-11-17',entries_text:'Fecha\tAna Pérez\tLuis Díaz\n01/10/2026\t2\t8'}],plan=D.plan(m,named),slide=plan.slides[7];assert.equal(slide.title,'Llamadas y avance por responsable');assert.deepEqual(slide.blocks[0].values[1],['Ana Pérez','2','5','12','10','2','20,0 %']);assert.deepEqual(slide.blocks[0].values[2],['Luis Díaz','8','1','2','4','0','0,0 %']);
 const text=JSON.stringify(plan.slides);assert(!/[·•]/.test(text));assert(!text.includes('Cartera'));assert(!text.includes('cartera'));assert.equal(plan.slides[0].subtitle,'17/11/2026');assert(!JSON.stringify(plan.slides[2]).includes('Ticket total ='));assert(!plan.slides.some(s=>'note' in s));
});
test('conversion cohort belongs to the first observed owner even after reassignment or reopening',()=>{
 const rows=[row('x','s','G - Sales Ready','2026-10-20',0),row('x','o','Closed Won','2026-10-20',100),row('x','n','B - Negotiate','2026-11-30',100)];rows[0].data.owner='Ana';rows[1].data.owner='Luis';rows[2].data.owner='Luis';
 const m=R.model(imports,rows,'n');assert.deepEqual(m.ownerConversion,[{owner:'Ana',worked:1,converted:0,rate:0}]);assert.equal(m.reopened,1);
});
test('presentation plan derives changing months and amounts from selection without future reports',()=>{
 const before=JSON.stringify(snapshots),m=R.model(imports,snapshots,'n'),plan=D.plan(m,calls);assert.equal(plan.slides.length,11);assert(!JSON.stringify(plan.slides).includes('Corte'));assert(!JSON.stringify(plan.slides).includes('Fecha de corte'));assert.equal(plan.calls.total,32);assert.equal(plan.asof,'2026-11-17');assert.equal(plan.slides[0].subtitle.includes('17/11/2026'),true);assert.match(JSON.stringify(plan),/250,00/);assert(!JSON.stringify(plan).includes('350,00'));
 const next=D.plan(R.model(imports,snapshots,'d'),calls);assert.equal(next.calls.total,35);assert.match(JSON.stringify(next),/diciembre de 2026/);assert.match(JSON.stringify(next),/350,00/);assert.equal(JSON.stringify(snapshots),before);
 const old=D.plan(R.model(imports,snapshots,'o'),calls);assert.equal(old.calls.available,false);assert.match(JSON.stringify(old),/Sin reporte utilizable/);
});
test('long histories paginate monthly evidence without dropping totals',()=>{
 const m=R.model(imports,snapshots,'n');m.updatedPeriods=Array.from({length:18},(_,i)=>({...m.updatedPeriods[0],month:new Date(Date.UTC(2025,5+i,1)).toISOString().slice(0,7)}));const plan=D.plan(m,[]);const pages=plan.slides.filter(s=>s.title.startsWith('Cierres por mes'));assert.equal(pages.length,4);assert.equal(pages.flatMap(s=>s.blocks[0].values).filter(r=>r[0]==='Total histórico').length,1);assert.equal(pages.flatMap(s=>s.blocks[0].values).filter(r=>r[0]==='Desde octubre 2026').length,1);
});
test('generated PPTX keeps brand template, escaped text, native tables, chart workbook and no speaker notes',async()=>{
 const m=R.model(imports,snapshots,'n'),bytes=await D.presentation(m,calls,template(),Zip,Excel,{outputType:'nodebuffer'}),z=await Zip.loadAsync(bytes),pres=await z.file('ppt/presentation.xml').async('string');assert.equal((pres.match(/<p:sldId\b/g)||[]).length,11);
 let xml='';for(let i=1;i<=11;i++){const s=await z.file('ppt/slides/slide'+i+'.xml').async('string');assert(!s.includes('{{'));assert.match(s.match(/<p:sld\b[^>]*>/)[0],/xmlns:a=/);xml+=s;const rel=await z.file('ppt/slides/_rels/slide'+i+'.xml.rels').async('string');assert(!rel.includes('/notesSlide'));}assert(xml.includes('<a:tbl>'));assert(xml.includes('&lt;Cuenta &amp; prueba&gt;'));assert(!xml.includes('<Cuenta & prueba>'));assert(xml.includes('Propietario &lt;Prueba&gt;'));assert(!xml.includes('48.611'));
 const chart=await z.file('ppt/charts/results-calls-8.xml').async('string');assert.match(chart,/<c:pt idx="0"><c:v>20<\/c:v>/);assert.match(chart,/<c:pt idx="1"><c:v>12<\/c:v>/);assert(!chart.includes('35'));
 const wb=new Excel.Workbook();await wb.xlsx.load(await z.file('ppt/embeddings/results-calls-8.xlsx').async('nodebuffer'));assert.equal(wb.getWorksheet('Datos').getCell('B2').value,20);assert.equal(wb.getWorksheet('Datos').getCell('B3').value,12);assert(z.file('ppt/media/image16.png'));
 assert.equal(Object.keys(z.files).filter(p=>/^ppt\/notes(?:Slides|Masters)\//.test(p)).length,0);assert(!(await z.file('[Content_Types].xml').async('string')).includes('/notesSlide'));
 for(const part of Object.keys(z.files).filter(p=>/^ppt\/(?:slides|slideMasters|slideLayouts)\/[^/]+\.xml$/.test(p))){const source=await z.file(part).async('string');for(const match of source.matchAll(/<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>/g))assert(!/[.·•]/.test(match[1]),part);}
});
test('results Excel includes date-specific prospecting only when available',async()=>{
 const m=R.model(imports,snapshots,'n'),p=P.model(calls,m.asof,m.first),bytes=await R.workbook(m,Excel,null,p),wb=new Excel.Workbook();await wb.xlsx.load(bytes);const sh=wb.getWorksheet('Prospección');assert(sh);assert.equal(sh.getCell('C8').value,20);assert.equal(sh.getCell('C9').value,12);assert(sh.getCell('A8').value instanceof Date);assert.equal(sh.getCell('E8').value,'Actualización pendiente de confirmar');
});
