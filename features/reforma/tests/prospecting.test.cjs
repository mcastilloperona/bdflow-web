const {test}=require('node:test'),assert=require('node:assert/strict'),P=require('../prospecting.js');
test('named Excel columns preserve blanks and validate totals through a save round trip',()=>{
 const text='Fecha\tAna Pérez\tLuis Díaz\tTotal general\n01/10/2026\t2\t\t2\n02/10/2026\t3\t7\t10';
 const rows=P.parse(text,'2026-10-02');assert.deepEqual(P.parse(P.serialize(rows),'2026-10-02'),rows);
 const m=P.model([{report_date:'2026-10-02',entries_text:P.serialize(rows)}],'2026-10-02','2026-10-01');assert.equal(m.total,12);assert.deepEqual(m.people.map(p=>[p.name,p.count,p.activeDays]),[['Luis Díaz',7,1],['Ana Pérez',5,2]]);
 assert.throws(()=>P.parse(text.replace('3\t7\t10','3\t7\t9'),'2026-10-02'),/total/);
 assert.throws(()=>P.parse('Fecha\tAna\tana\n01/10/2026\t1\t2','2026-10-02'),/únicos/);
 assert.equal(P.conversion('Ana Pérez',[{owner:'Ana María PEREZ',worked:10}]).worked,10);assert.equal(P.conversion('Ana Pérez',[{owner:'Ana PEREZ'},{owner:'Ana María PEREZ'}]),null);
});
test('prospecting validates calendar dates, whole counts, cutoff and duplicate rows',()=>{
 assert.deepEqual(P.parse('Fecha\tTotal\n01/10/2026\t0\n2026-10-02;12','2026-10-02'),[{date:'2026-10-01',count:0},{date:'2026-10-02',count:12}]);
 for(const text of ['31/02/2026 1','2026-10-03 2','2026-10-01 -1','2026-10-01 1.5','2026-10-01 1\n2026-10-01 2','2026-10-01 2 3','2026-10-01 100001'])assert.throws(()=>P.parse(text,'2026-10-02'));
});
test('reports use their own historical cutoff, exclude dates before initiative and do not sum successive snapshots',()=>{
 const reports=[{id:'a',report_date:'2026-11-01',entries_text:'2026-09-28 1\n2026-10-01 10',record_complete:false},{id:'b',report_date:'2026-11-18',entries_text:'2026-10-01 15\n2026-11-18 20',record_complete:true}];
 const p=P.model(reports,'2026-11-17','2026-09-29');assert.equal(p.total,10);assert.equal(p.asof,'2026-11-01');assert.equal(p.complete,false);assert.deepEqual(p.months,[{month:'2026-10',count:10}]);
 assert.equal(P.model(reports,'2026-11-18','2026-09-29').total,35);assert.equal(P.model(reports,'2026-10-31','2026-09-29').available,false);
 const bad=P.model([{report_date:'2026-11-01',entries_text:'bad'}],'2026-11-17','2026-09-29');assert.equal(bad.available,false);assert(bad.error);
});
