const {test}=require('node:test'),assert=require('node:assert/strict'),R=require('../results.js'),C=require('../core.js'),D=require('../results-presentation.js');
const imports=[{id:'s',report_date:'2026-09-29'},{id:'o',report_date:'2026-10-29'},{id:'n',report_date:'2026-11-18'},{id:'f',report_date:'2026-12-01'}];
const row=(op,imp,stage,close,amount=100,owner='Ana Pérez')=>({id:op+imp,import_id:imp,operation_id:op,data:{account:op,opportunity:op,owner,stage,close_date:close,amount,amount_basis:'saas',target_plan:C.TARGET}});
test('a win first received in November is an October closure, including direct wins and later corrections',()=>{
 const rows=[row('late','s','G - Sales Ready','2026-10-31',0),row('late','o','B - Negotiate','2026-10-31'),row('late','n','Closed Won','2026-10-31',200),row('direct','n','Closed Won','2026-10-30',50),row('late','f','Closed Won','2026-10-30',250)];
 const m=R.model(imports,rows,'n'),oct=m.updatedPeriods.find(p=>p.month==='2026-10'),nov=m.updatedPeriods.find(p=>p.month==='2026-11');
 assert.equal(oct.count,2);assert.equal(oct.amount,250);assert.equal(oct.conversions,1);assert.equal(oct.timings.sample,1);assert.equal(oct.timings.mean,32);assert.equal(nov.count,0);assert.equal(nov.amount,0);assert.equal(nov.conversions,0);assert.equal(m.events[0].firstWon,'2026-11-18');assert.equal(m.events[0].close_date,'2026-10-31');assert.equal(m.periods.find(p=>p.month==='2026-10').convertedAmount,200);assert.equal(m.periods.find(p=>p.month==='2026-11').conversions,0);
 const slide=D.plan(m,[]).slides[5],values=slide.blocks[0].values;assert.deepEqual(values[1],['octubre de 2026','2','250,00 €','1','32,0','32,0']);assert.deepEqual(values[2],['noviembre de 2026','0','0,00 €','0','No disponible','No disponible']);
});
test('close dates before first work never create negative elapsed time and future closes are excluded',()=>{
 const m=R.model(imports,[row('past','o','B - Negotiate','2026-10-20'),row('past','n','Closed Won','2026-10-20'),row('future','s','G - Sales Ready','2026-12-01'),row('future','n','Closed Won','2026-12-01')],'n');
 assert.equal(m.currentWon.count,1);assert.equal(m.events[0].days,null);assert.equal(m.updatedPeriods.find(p=>p.month==='2026-10').timings.sample,0);assert.equal(m.cumulative.timings.sample,0);
});
test('net stage progress ignores oscillations, new arrivals, lost states and future snapshots',()=>{
 const rows=[row('a','s','G - Sales Ready','2026-11-30'),row('a','o','B - Negotiate','2026-11-30'),row('a','n','D - Draw up the solution','2026-11-30'),row('a','f','A - Sign','2026-12-30'),row('new','n','A - Sign','2026-11-30'),row('lost','s','G - Sales Ready','2026-11-30'),row('lost','n','Closed Lost','2026-11-30')];
 const m=R.model(imports,rows,'n');assert.deepEqual(m.ownerProgress,[{owner:'Ana Pérez',worked:3,advanced:1,steps:3,rate:1/3}]);
});
test('ownership transfers reset stage comparison and winning owner receives the closure',()=>{
 const rows=[row('a','s','G - Sales Ready','2026-10-20',0,'Ana Pérez'),row('a','o','B - Negotiate','2026-10-20',100,'Luis Díaz'),row('a','n','Closed Won','2026-10-20',100,'Luis Díaz')],m=R.model(imports,rows,'n');
 assert.deepEqual(m.ownerProgress.map(g=>[g.owner,g.advanced,g.steps]),[['Ana Pérez',0,0],['Luis Díaz',1,2]]);assert.equal(m.ownerConversion.find(g=>g.owner==='Ana Pérez').converted,0);assert.equal(m.ownerConversion.find(g=>g.owner==='Luis Díaz').converted,1);
 const late=R.model(imports,[row('direct','n','Closed Won','2026-10-30',50,'Luis Díaz')],'n');assert.equal(late.ownerConversion[0].worked,1);assert.equal(late.ownerConversion[0].converted,1);assert.equal(late.ownerProgress[0].advanced,0);
});
test('requested slide comments are absent and both slide 4 tables share row geometry',()=>{
 const m=R.model(imports,[row('a','s','G - Sales Ready','2026-10-20'),row('a','n','Closed Won','2026-10-20')],'n'),plan=D.plan(m,[]),left=plan.slides[3].blocks.filter(b=>b.kind==='table')[0],right=plan.slides[3].blocks.filter(b=>b.kind==='table')[1];
 assert.equal(left.y,right.y);assert.equal(left.h,right.h);assert.equal(left.values.length,right.values.length);assert.equal(plan.slides[3].blocks.filter(b=>b.kind==='text').length,2);assert.equal(plan.slides[4].blocks.filter(b=>b.kind==='text').length,0);assert.equal(plan.slides[5].blocks.filter(b=>b.kind==='text').length,0);assert(!JSON.stringify(plan.slides[6]).includes('Describe el ritmo observado'));
});
