const {test}=require('node:test'),assert=require('node:assert/strict'),R=require('../results.js'),D=require('../results-presentation.js'),C=require('../core.js');
const imports=[{id:'s',report_date:'2026-09-29'},{id:'o',report_date:'2026-10-20'},{id:'n',report_date:'2026-11-18'},{id:'f',report_date:'2026-12-05'}];
const row=(id,imp,stage,owner='Ana Pérez',close='2026-10-31')=>({id:id+imp,operation_id:id,import_id:imp,data:{account:id,opportunity:id,owner,stage,close_date:close,amount:100,amount_basis:'saas',target_plan:C.TARGET}});
test('stage changes deduplicate repeated crossings and exclude reassignment, future imports and direct wins',()=>{
 const rows=[row('x','s','G - Sales Ready'),row('x','o','B - Negotiate'),row('x','n','G - Sales Ready'),row('x','f','B - Negotiate'),row('handoff','s','G - Sales Ready'),row('handoff','o','B - Negotiate','Luis Díaz'),row('handoff','n','Closed Won','Luis Díaz'),row('direct','n','Closed Won')];
 const m=R.model(imports,rows,'n');assert.equal(m.stageTransitions.length,2);const g=m.stageTransitions.find(t=>t.from==='G - Sales Ready');assert.deepEqual(g.owners,[{owner:'Ana Pérez',operationIds:['x'],count:1}]);assert.deepEqual(m.stageTransitions.find(t=>t.to==='Closed Won').owners,[{owner:'Luis Díaz',operationIds:['handoff'],count:1}]);
 const future=R.model(imports,rows,'f');assert.equal(future.stageTransitions.find(t=>t.from==='G - Sales Ready').owners[0].count,1);
});
test('export pivots actual callers dynamically and keeps the approved labels and clean slide comments',()=>{
 const rows=[row('a','s','G - Sales Ready'),row('a','o','B - Negotiate'),row('b','s','A - Sign','Luis Díaz'),row('b','n','Closed Won','Luis Díaz')],m=R.model(imports,rows,'n'),reports=[{report_date:'2026-11-17',entries_text:'Fecha\tAna Pérez\tLuis Díaz\n01/10/2026\t4\t7'}],plan=D.plan(m,reports),slide=plan.slides.find(s=>s.title==='Avance de oportunidades por etapa'),values=slide.blocks[0].values;
 assert.deepEqual(values[0],['Etapa de origen','Etapa de destino','Ana','Luis','Total']);assert.deepEqual(values[1],['G - Sales Ready','B - Negotiate','1','0','1']);assert.deepEqual(values[2],['A - Sign','Closed Won','0','1','1']);assert.equal(slide.blocks.filter(b=>b.kind==='text').length,0);assert(!JSON.stringify(plan).includes('Etapas netas'));
 assert.equal(plan.slides.find(s=>s.title==='Prioridades para el cierre del mes').blocks.filter(b=>b.kind==='text').length,0);assert.deepEqual(plan.slides[5].blocks[0].values[0],['Mes de cierre','Ganadas','SaaS Booking','Tiempo medio hasta cierre (días)']);
});
test('timing footnote follows the selected data rather than a fixed company or month',()=>{
 const rows=[row('Untracked Example','n','Closed Won')],m=R.model(imports,rows,'n'),plan=D.plan(m,[]);const note=plan.slides[5].blocks.find(b=>b.kind==='text');assert(note.text.includes('Untracked Example'));assert(note.text.includes('31/10/2026'));assert(!JSON.stringify(plan.slides[5]).includes('Con tiempo'));assert(!JSON.stringify(plan.slides[5]).includes('Mediana'));
 const old=D.plan(R.model(imports,rows,'o'),[]);assert(!JSON.stringify(old.slides[5]).includes('Untracked Example'));
});
