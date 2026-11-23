import assert from 'node:assert/strict';
import {parseEvents,parseStamp,duration} from './dist/core.mjs';
import {parsePresence,parseSchedule,buildShiftReport,aggregateShops,reportCoverage} from './dist/plan.mjs';
const O='Неконтролируемая территория',H='Территория холдинга',W='Цех №2';
const header=['Фамилия','Имя','Отчество','Табельный номер','Должность','Подразделение','Дата','Вход','Выход','Присутствие','График работы'];
const p=(day,entry='',exit='',time='00:00',surname='Заводской',id='0002',dept='Участок')=>[surname,'Петр','Петрович',id,'Рабочий',dept,`${String(day).padStart(2,'0')}-11-2026`,entry,exit,time,'С 07:00 до 19:00'];
const input=[header,p(1,'18:50:00','','05:10'),p(1,'18:50:00','','05:10'),p(2,'18:55:00','07:10:00','12:15'),p(3,'','07:15:00','07:15'),p(4),p(5,'08:05:00','18:50:00','10:45'),p(6,'19:00:00','','05:00'),p(7,'','','08:00'),p(7,'06:00:00','20:00:00','14:00','Посторонний','999'),p(4,'06:00:00','20:00:00','14:00','Сменщик','998')];
input.rowSources=input.map((_,i)=>({sheetName:'Сводка',rowNumber:10+i}));
const presence=parsePresence(input,'summary.xlsx');
const person=presence.people.find(p=>p.id==='2');
assert.equal(person.events.length,7,'duplicate rows must not duplicate boundaries');
assert.equal(person.events[0].rowNumber,11);assert.equal(person.events[0].sheetName,'Сводка');
assert.equal(person.events.find(e=>e.ts===parseStamp('02-11-2026 07:10:00')).to,O,'morning exit stays on its calendar date');
assert.equal(person.events.find(e=>e.ts===parseStamp('02-11-2026 18:55:00')).from,O,'evening entry on the same row is a separate boundary');
const planRows=[['Ноябрь 2026'],['ФИО','Таб. №','1','2','3','4','5','6','7'],['Цехов И.И.','1','11','11','11','11','11','11','11'],['Заводской П.П.','2','11','11','в','11','11','11','11']];
planRows.cellStyles=[[],[],[],[{}, {},{fillColor:'808080'},{fillColor:'808080'},{fillColor:'FFFFFF'},{fillColor:'FFFFFF'},{fillColor:'FFFFFF'},{fillColor:'808080'},{fillColor:'FFFFFF'}]];
const plan=parseSchedule(planRows,'plan.xlsx');
const detailedHeader=['Фамилия','Имя','Отчество','Табельный номер','Дата','Выход из','Вход в','Подразделение'];
const readDetailed = rows => parseEvents([detailedHeader, ...rows.map(([ts,from,to])=>['Цехов','Иван','Иванович','1',ts,from,to,'Цех №2'])], 'detail.xlsx').events;
const detailed=readDetailed([
 ['01-11-2026 06:40:00',O,H],['01-11-2026 06:50:00',H,W],['01-11-2026 19:00:00',W,H],['01-11-2026 19:10:00',H,O],
 ['02-11-2026 06:40:00',O,H],['02-11-2026 06:50:00',H,W],['02-11-2026 19:00:00',W,H],['02-11-2026 19:10:00',H,O],
]);
const report=buildShiftReport(detailed,plan,presence,'2026-11-01','2026-11-30');
const factory=report.rows.filter(r=>r.id==='2'),shops=report.rows.filter(r=>r.id==='1');
assert(factory.length>0);
const unplanned = report.rows.filter(r => r.id === '999');
assert.equal(unplanned.length, 1, 'include actual marks for employees who appear only in the summary');
assert.equal(unplanned[0].accounting, 'factory');
assert.equal(unplanned[0].status, 'По норме K');
assert(shops.every(r=>r.date<='2026-11-02'),'later summary data must not create future shop rows');
const first=factory.find(r=>r.date==='2026-11-01');
assert.equal(first.accounting,'factory');assert.equal(first.factSource,'Сводка PERCo');assert.equal(duration(first.span),'12:20');assert.equal(first.last.ts,parseStamp('02-11-2026 07:10:00'));assert.equal(first.level,'green');
assert.equal(duration(factory.find(r=>r.date==='2026-11-02').span),'12:20');
assert.equal(factory.find(r=>r.date==='2026-11-04').level,'red','known missing shift from summary is red');
const late=factory.find(r=>r.date==='2026-11-05');assert(late.comments.some(c=>c.includes('Поздний вход на завод на 65 мин')));assert(late.comments.some(c=>c.includes('Ранний выход с завода на 10 мин')));
const open=factory.find(r=>r.date==='2026-11-06');assert.equal(open.last,null);assert.equal(open.span,null,'never replace missing exit with midnight or reported calendar-day duration');
const positive=factory.find(r=>r.date==='2026-11-07');assert.equal(positive.level,'gray');assert(positive.comments.some(c=>c.includes('время присутствия')));
assert(factory.every(r=>!r.segments.length&&!r.shopGross&&!r.lunchShop),'summary boundaries must not create shop intervals or shop lunch');
assert(aggregateShops(report,'day').rows.every(r=>r.id==='1'));
assert.equal(reportCoverage(detailed,presence).end,parseStamp('07-11-2026 20:00:00'),'summary marks extend the report');
assert(buildShiftReport([],plan,presence,'2026-11-01','2026-11-30').rows.some(r=>r.id==='2'&&r.first),'summary plus monthly plan works without detailed passes');
const summaryOnly = buildShiftReport([],null,presence,'2026-11-01','2026-11-30').rows;
assert(summaryOnly.some(r=>r.id==='2' && r.first),'include summary marks with no monthly roster');
assert(!summaryOnly.some(r=>r.id==='2' && r.date==='2026-11-04'),'a blank K day with no assigned shift is not a no-show');
const overlap=parsePresence([header,...[p(1,'05:00','21:00','16:00','Цехов','1','Цех №2')].map(row=>{row[1]='Иван';row[2]='Иванович';return row;})]);
const unchanged=buildShiftReport(detailed,plan,overlap,'2026-11-01','2026-11-01').rows.find(r=>r.id==='1');
assert.equal(unchanged.first.ts,parseStamp('01-11-2026 06:50:00'));assert.equal(unchanged.accounting,'shop','detail wins for a person present in both files');
assert(!unchanged.comments.some(c=>c.includes('Проверьте табельные номера')));

const mismatchedPlan = parseSchedule([
  ...planRows.slice(0,2),
  ['Цехов И. И.','42','11'],
  ['Цехов И. П.','43','11'],
], 'plan.xlsx');
const numberRows = buildShiftReport(detailed,mismatchedPlan,overlap,'2026-11-01','2026-11-01').rows;
const plannedIdentity = numberRows.find(r=>r.id==='42');
const observedIdentity = numberRows.find(r=>r.id==='1');
for (const row of [plannedIdentity, observedIdentity]) {
  assert(row.comments.some(c=>c.includes('в месячном графике — № 42') && c.includes('в выгрузках проходов — № 1')));
  assert(row.needsReview);
  assert(row.comments.includes('Нужно ручное рассмотрение.'));
}
assert.equal(plannedIdentity.status,'Недостаточно данных');
assert.equal(plannedIdentity.first,null,'a similar name must not receive another number’s timestamps');
assert.equal(observedIdentity.first.ts,unchanged.first.ts);
assert.equal(observedIdentity.last.ts,unchanged.last.ts);
assert(!numberRows.find(r=>r.id==='43').comments.some(c=>c.includes('Проверьте табельные номера')));

const ambiguous=parsePresence([header,p(1,'06:00','19:00'),p(1,'06:00','19:00','13:00','Заводской','2','Другое подразделение')]);
const ambiguousRows=buildShiftReport([],plan,ambiguous,'2026-11-01','2026-11-01').rows.filter(r=>r.id==='2');
assert.equal(ambiguousRows.find(r=>r.factSource==='Только месячный график').level,'gray');
assert(ambiguousRows.every(r=>r.level!=='red'));
const invalid=parsePresence([header,p(1,'25:10','19:00')]);
assert(invalid.people[0].days['2026-11-01'][0].issues.length>0);
const fixedPresence = parsePresence([header, ...input.slice(1).map(row => [...row.slice(0,10), 'С 07:00 до 16:00'])], 'summary.xlsx');
const conflicted = buildShiftReport(detailed,plan,fixedPresence,'2026-11-01','2026-11-07').rows.find(r=>r.id==='2'&&r.date==='2026-11-04');
assert.equal(conflicted.status,'Не вышел на смену','different hours do not erase a documented no-show');
assert(conflicted.comments.some(c=>c.includes('07:00–19:00, 12 ч') && c.includes('07:00–16:00, 9 ч')));
assert(conflicted.comments.includes('Нужно ручное рассмотрение.'));
const transfers = parsePresence([header,p(1,'06:00','19:05','13:05','Заводской','2','Цех 1'),p(2,'06:00','19:05','13:05','Заводской','2','Цех 2')]);
const transferRows=buildShiftReport([],null,transfers,'2026-11-01','2026-11-02').rows;
assert.equal(transferRows.length,2);
assert.equal(new Set(transferRows.map(r=>r.employeeKey)).size,1,'non-overlapping department assignments use the same employee key');
const exitOnly = parsePresence([header,p(1,'','07:10','07:10')]);
const firstDay = buildShiftReport([],null,exitOnly,'2026-11-01','2026-11-01').rows[0];
assert.equal(firstDay.date,'2026-11-01');
assert.equal(firstDay.first,null);
assert(firstDay.last && firstDay.level!=='red','keep the known exit, with a missing entry');

const rowForShopEmployee = (day, entry, exit) => {
  const row = p(day,entry,exit,'00:00','Цехов','1','Цех №2');
  return [row[0],'Иван','Иванович',...row.slice(3)];
};
const secondary = parsePresence([header,
  rowForShopEmployee(2,'05:00','21:00'),
  rowForShopEmployee(3,'06:57','19:12'),
  rowForShopEmployee(4,'06:45',''),
  rowForShopEmployee(5,'18:40',''),
  rowForShopEmployee(6,'','07:15'),
],'summary.xlsx');
const mixed = buildShiftReport(detailed,null,secondary,'2026-11-02','2026-11-06').rows;
assert(mixed.every(r=>r.accounting==='shop'),'gate fallback stays on the detailed sheet');
assert.equal(new Set(mixed.map(r=>r.employeeKey)).size,1);
assert.equal(mixed.filter(r=>r.date==='2026-11-02').length,1,'shop and factory data never produce a duplicate row');
assert.equal(mixed[0].first.ts,parseStamp('02-11-2026 06:50:00'),'prefer real shop boundaries when available');
const fallback = mixed.find(r=>r.date==='2026-11-03');
assert.equal(fallback.first.ts,parseStamp('03-11-2026 06:57:00'));
assert.equal(fallback.last.ts,parseStamp('03-11-2026 19:12:00'));
assert.equal(fallback.factSource,'Сводка PERCo · проходная');
assert.equal(fallback.status,'Был на заводе');
assert(fallback.needsReview && !fallback.shopGross);
assert.equal(mixed.find(r=>r.date==='2026-11-04').span,null,'do not use a later exit for a previous unclosed visit');
const overnightFallback=mixed.find(r=>r.date==='2026-11-05');
assert.equal(duration(overnightFallback.span),'12:35');
assert.equal(mixed.filter(r=>r.date==='2026-11-06').length,0,'morning exit belongs to the preceding night');

const detailedGateEvents = readDetailed([
  ['03-11-2026 06:55:00',O,'Гостевой'], ['03-11-2026 12:00:00',H,O],
  ['03-11-2026 12:30:00',O,H], ['03-11-2026 19:15:00',H,O],
  ['04-11-2026 18:55:00',O,H], ['05-11-2026 07:15:00',H,O],
  ['06-11-2026 06:55:00',O,H],
]);
const detailGateRows = buildShiftReport(detailedGateEvents,null,null,'2026-11-03','2026-11-06').rows;
assert.equal(detailGateRows.length,3);
const [territoryDay, territoryNight, openTerritory] = detailGateRows;
assert.equal(territoryDay.status,'Был на заводе');
assert.equal(territoryDay.accounting,'shop','keep the original employee sheet');
assert.equal(territoryDay.factSource,'Подробные проходы · проходная');
assert.equal(territoryDay.first.fileName,'detail.xlsx');
assert.equal(duration(territoryDay.span),'12:20');
assert.equal(duration(territoryDay.gaps[0]),'0:30');
assert.equal(territoryDay.shopGross,0);
assert(territoryDay.needsReview && territoryDay.comments.some(c=>c.includes('из подробной выгрузки')));
const dayOffPlan = parseSchedule([['Ноябрь 2026'],['ФИО','Таб. №','3'],['Цехов И.И.','1','в']]);
const offScheduleGate = buildShiftReport(detailedGateEvents,dayOffPlan,null,'2026-11-03','2026-11-03').rows[0];
assert.equal(offScheduleGate.status,'Выход вне графика');
assert.equal(offScheduleGate.level,'red');
assert.equal(offScheduleGate.first.ts,territoryDay.first.ts);
assert.equal(territoryNight.date,'2026-11-04');
assert.equal(territoryNight.last.ts,parseStamp('05-11-2026 07:15:00'));
assert.equal(openTerritory.status,'Был на заводе');
assert.equal(openTerritory.last,null);
assert.equal(openTerritory.span,null);
const isolatedDetailedExit = buildShiftReport(readDetailed([['05-11-2026 07:15:00',H,O]]),null,null,'2026-11-05','2026-11-05').rows[0];
assert.equal(isolatedDetailedExit.status,'Зафиксирован выход с завода');
assert.equal(isolatedDetailedExit.first,null);
assert.equal(isolatedDetailedExit.span,null);

const partialSummary = parsePresence([header,rowForShopEmployee(3,'06:55','')],'summary.xlsx');
const completeGate = buildShiftReport(detailedGateEvents,null,partialSummary,'2026-11-03','2026-11-03').rows[0];
assert.equal(completeGate.factSource,'Подробные проходы · проходная');
assert.equal(completeGate.last.ts,territoryDay.last.ts,'a complete detailed visit wins over a partial summary');
const partialShop = buildShiftReport(readDetailed([
  ['03-11-2026 06:55:00',O,H], ['03-11-2026 07:01:00',H,W], ['03-11-2026 19:15:00',H,O],
]),null,partialSummary,'2026-11-03','2026-11-03').rows[0];
assert.equal(partialShop.first.ts,parseStamp('03-11-2026 07:01:00'));
assert.equal(partialShop.last,null,'do not use the gate exit as a shop exit');
assert(partialShop.comments.some(c=>c.includes('подтверждают присутствие на заводе')));

const summaryGap = parsePresence([header,p(1,'07:00','19:15'),p(2),p(3,'07:00','19:15')],'summary-gap.xlsx');
const gapRow=buildShiftReport([],plan,summaryGap,'2026-11-02','2026-11-02').rows.find(r=>r.id==='2');
assert.equal(gapRow.status,'Недостаточно данных');
assert(gapRow.comments.some(c=>c.includes('нет ни одной отметки за 02.11.2026')));
assert(gapRow.needsReview);
console.log('Summary attendance checks passed: both sources, roster scope, overnight boundaries, duplicates, priority, source cutoff, missing/invalid marks and no invented shop time.');
