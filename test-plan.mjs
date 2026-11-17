import assert from 'node:assert/strict';
import { parseEvents, duration } from './dist/core.mjs';
import { parseSchedule, parsePresence, buildShiftReport, aggregateShops } from './dist/plan.mjs';

const headers = ['Фамилия','Имя','Отчество','Табельный номер','Дата','Выход из','Вход в','Подразделение'];
const O='Неконтролируемая территория', H='Территория холдинга', W='Цех №2';
const event = (ts,from,to,id='0001',surname='Тест',dept='Цех №2') => [surname,'Иван','Иванович',id,ts,from,to,dept];
const events = rows => parseEvents([headers,...rows]).events;
function schedule(code='11',color='FFFFFF',dates=['1'],month='ноябрь 2026') {
  const rows=[[month],['ФИО','Таб. №',...dates],['Тест И.И.','1',...dates.map(()=>code)]];
  rows.cellStyles=[[],[],[{}, {},...dates.map(()=>({fillColor:color}))]];
  return parseSchedule(rows,'plan.xlsx');
}
function presence(template, surname='Тест',dept='Цех №2') {
  return parsePresence([['Фамилия','Имя','Отчество','Табельный номер','Подразделение','Дата','График работы'],[surname,'Иван','Иванович','1',dept,'01-11-2026',template]],'presence.xlsx');
}
const complete=events([
  event('01-11-2026 06:30:00',O,H),event('01-11-2026 07:05:00',H,W),event('01-11-2026 18:55:00',W,H),event('01-11-2026 19:20:00',H,O),
]);
let r=buildShiftReport(complete,schedule(),presence('С 07:00 до 19:00'),'2026-11-01','2026-11-01').rows[0];
assert.equal(duration(r.span),'11:50');
assert.equal(r.level,'red');
assert(r.comments.includes('Поздний вход в цех: 5 мин'));
assert(r.comments.includes('Ранний выход из цеха: 5 мин'));
assert.equal(duration(r.lunchShop),'1:00');

const night=events([event('30-11-2026 18:30:00',O,H),event('30-11-2026 18:50:00',H,W),event('01-12-2026 07:00:16',W,H),event('01-12-2026 07:20:00',H,O)]);
r=buildShiftReport(night,schedule('11','808080',['30']),presence('С 07:00 до 19:00'),'2026-11-01','2026-11-30').rows[0];
assert.equal(r.date,'2026-11-30');assert.equal(r.status,'Вышел в свою смену');assert.equal(duration(r.span),'12:10');
assert.equal(buildShiftReport(night,schedule('11','808080',['30']),null,'2026-12-01','2026-12-31').rows.length,0);

const reentries=events([event('01-11-2026 06:40:00',O,H),event('01-11-2026 06:50:00',H,W),event('01-11-2026 12:00:00',W,H),event('01-11-2026 12:10:00',H,O),event('01-11-2026 13:10:00',O,H),event('01-11-2026 13:20:00',H,W),event('01-11-2026 19:00:00',W,H),event('01-11-2026 19:10:00',H,O)]);
let report=buildShiftReport(reentries,schedule(),presence('С 07:00 до 19:00'),'2026-11-01','2026-11-01');
r=report.rows[0];assert.equal(report.rows.length,1);assert.equal(duration(r.span),'12:10');assert.equal(duration(r.shopGross),'10:50');assert.equal(duration(r.lunchShop),'1:00');assert.equal(duration(r.gaps[0]),'1:00');
assert.equal(duration(aggregateShops(report,'month').rows[0].lunchShop),'1:00');

const partial=complete.slice(0,2);
r=buildShiftReport(partial,schedule(),presence('С 07:00 до 19:00'),'2026-11-01','2026-11-01').rows[0];
assert.equal(r.last,null);assert(!r.comments.some(c=>c.startsWith('Ранний выход')));assert.equal(r.status,'Недостаточно данных');

r=buildShiftReport(complete,schedule('11','808080'),presence('С 07:00 до 16:00'),'2026-11-01','2026-11-01').rows[0];
assert.equal(new Date(r.plan.start).getHours(),7);assert.equal(new Date(r.plan.end).getHours(),16);assert(r.comments.includes('Ночная ячейка при дневном графике K'));
r=buildShiftReport(complete,schedule('8'),presence('С 08:00 до 17:00'),'2026-11-01','2026-11-01').rows[0];
assert.equal(new Date(r.plan.start).getHours(),8);assert.equal(new Date(r.plan.end).getHours(),17);

const foreign=presence('С 08:00 до 17:00','Чужой','Другой завод');
r=buildShiftReport(complete,schedule(),foreign,'2026-11-01','2026-11-01').rows[0];assert.equal(r.profile,null);assert.equal(new Date(r.plan.end).getHours(),19);
const normalized=events([event('01-11-2026 06:40:00',O,H,'00 01'),event('01-11-2026 06:50:00',H,W,'0001'),event('01-11-2026 19:00:00',W,H,'1'),event('01-11-2026 19:10:00',H,O,'0001')]);
assert.equal(buildShiftReport(normalized,schedule(),null,'2026-11-01','2026-11-01').rows.length,1);

const dayOff=buildShiftReport(complete,schedule('в'),null,'2026-11-01','2026-11-01').rows[0];assert.equal(dayOff.status,'Выход вне графика');
const leadingExit=events([event('30-11-2026 07:00:00',W,H),event('30-11-2026 07:15:00',H,O),...night.map(e=>event(new Date(e.ts).toLocaleString('ru-RU').replace(',',''),e.from,e.to))]);
r=buildShiftReport(leadingExit,schedule('11','808080',['30']),presence('С 07:00 до 19:00'),'2026-11-30','2026-11-30').rows[0];
assert.equal(r.status,'Вышел в свою смену');assert.equal(r.level,'green');
console.log('Plan/fact checks passed: workshop boundaries, day/night, reentries, lunch, partial coverage, fixed daytime templates, employee identity.');
