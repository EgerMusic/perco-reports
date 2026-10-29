import assert from 'node:assert/strict';
import { parseEvents, analyze, buildReport, duration, issueKey } from './dist/core.mjs';

const rows = [
  ['Отчёт'],
  ['Фамилия','Имя','Отчество','Табельный номер','Дата','Выход из','Вход в','Карта №'],
  ['Тест','Иван','','001','30-11-2026 18:40:00','Неконтролируемая территория','Гостевой','1'],
  ['Тест','Иван','','001','30-11-2026 18:55:00','Территория холдинга','Цех А','1'],
  ['Тест','Иван','','001','01-12-2026 06:00:00','Цех Б','Территория холдинга','1'],
  ['Тест','Иван','','001','01-12-2026 07:20:00','Гостевой','Неконтролируемая территория','1'],
];
const events = parseEvents(rows).events;
assert.equal(events.length, 4);
const initial = analyze(events);
assert.equal(initial.sessions.length, 1);
assert.equal(initial.issues.length, 1);
assert.equal(duration(initial.sessions[0].end - initial.sessions[0].start), '12:40');
assert.equal(initial.sessions[0].entry.rowNumber, 3);
assert.equal(initial.sessions[0].exit.rowNumber, 6);
const november = buildReport(initial, '2026-11-01', '2026-11-30');
const december = buildReport(initial, '2026-12-01', '2026-12-31');
assert.equal(duration(november.rows[0].plant), '5:20');
assert.equal(duration(december.rows[0].plant), '7:20');
assert.equal(december.rows[0].visits.length, 1);
assert.equal(december.rows[0].visits[0].start, events[0].ts);
assert.equal(duration(initial.segments.find(x => x.kind === 'unknown').end - initial.segments.find(x => x.kind === 'unknown').start), '11:05');
const cut = new Date('2026-12-01T01:00:00').getTime();
const fixed = analyze(events, { [issueKey(events[1], events[2])]: cut });
assert.equal(fixed.issues.length, 0);
assert.equal(duration(fixed.segments.filter(x => x.kind === 'shop').reduce((n, x) => n + x.end - x.start, 0)), '11:05');
assert.equal(duration(buildReport(fixed, '2026-11-30', '2026-12-01').rows.reduce((n, x) => n + x.plant, 0)), '12:40');
const split = buildReport(fixed, '2026-11-30', '2026-12-01');
assert.equal(duration(split.rows.reduce((n, x) => n + x.lunchShop, 0)), '1:00');
assert.equal(duration(buildReport(fixed, '2026-12-01', '2026-12-31').rows[0].lunchShop), '1:00');
console.log('Core checks passed');
