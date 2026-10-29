import { parseEvents, dedupeEvents, analyze, buildReport, duration, localDate, periodKey } from './core.mjs';
import { readXlsx } from './xlsx.mjs';
import { downloadWorkbook, excelStyle } from './export.mjs';

const $ = id => document.getElementById(id);
const collator = new Intl.Collator('ru', { sensitivity: 'base', numeric: true });
const shopColors = ['#eb8135','#2c8874','#4d7ea0','#9b755e','#836eac','#60926b','#ba786a','#7899a8'];
let rawEvents = [], files = [], analysis = null, report = null, exportSheets = [];
const esc = value => String(value ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);
const stamp = ms => new Date(ms).toLocaleString('ru-RU', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
const source = e => e ? `${e.fileName || 'Файл'} · строка ${e.rowNumber}` : '—';
const notify = message => { $('notice').textContent = message; };
const periodDate = period => period.startsWith('Неделя с ') ? period.slice(10) : period.length === 7 ? `${period}-01` : period;
const sortValue = r => r.ts ?? new Date(`${periodDate(r.period)}T00:00:00`).getTime();
const bySort = (a, b) => {
  const mode = $('sort').value;
  if (mode === 'name') return collator.compare(a.name, b.name) || sortValue(a) - sortValue(b) || collator.compare(a.id, b.id);
  return (mode === 'date-desc' ? -1 : 1) * (sortValue(a) - sortValue(b)) || collator.compare(a.name, b.name) || collator.compare(a.id, b.id);
};
const selectedBounds = () => [new Date(`${$('from-date').value}T00:00:00`).getTime(), new Date(`${$('to-date').value}T00:00:00`).getTime() + 86400000];

function shiftOf(ts) {
  const entered = new Date(ts);
  const hour = entered.getHours() + entered.getMinutes() / 60;
  const night = hour >= 13 || hour < 1;
  const start = new Date(entered);
  if (night && hour < 1) start.setDate(start.getDate() - 1);
  start.setHours(night ? 19 : 7, 0, 0, 0);
  return { start: start.getTime(), kind: night ? 'Ночная' : 'Дневная' };
}

function groupSessionsByShift(sessions) {
  const byEmployee = new Map();
  const groups = [];
  for (const session of [...sessions].sort((a, b) => a.start - b.start)) {
    const previous = byEmployee.get(session.id);
    const shift = previous && session.start <= previous.start + 12 * 3600000 + 10 * 60000
      ? previous : { id: session.id, name: session.name, ...shiftOf(session.start), first: session.start, last: session.end, visits: [] };
    shift.first = Math.min(shift.first, session.start);
    shift.last = Math.max(shift.last, session.end);
    shift.visits.push(session);
    byEmployee.set(session.id, shift);
    if (shift !== previous) groups.push(shift);
  }
  return groups;
}

function scheduleNotes(shifts) {
  const [from, to] = selectedBounds();
  const notes = new Map();
  for (const shift of shifts) {
    if (shift.start < from || shift.start >= to) continue;
    const entryLimit = shift.start - 10 * 60000;
    const exitLimit = shift.start + 12 * 3600000 + 10 * 60000;
    const flags = [];
    if (shift.first > entryLimit) flags.push(`Поздний вход ${stamp(shift.first)}`);
    if (shift.last < exitLimit) flags.push(`Ранний выход ${stamp(shift.last)}`);
    const gaps = shift.visits.slice(1).map((visit, index) => Math.max(0, visit.start - shift.visits[index].end)).filter(Boolean);
    const gapText = gaps.length === 1 ? `Выходил с завода на ${duration(gaps[0])}` : gaps.length > 1 ? `Выходил с завода ${gaps.length} раза, всего ${duration(gaps.reduce((sum, gap) => sum + gap, 0))}` : '';
    const period = periodKey(shift.start, $('group').value);
    const key = `${shift.id}|${period}`;
    if (!notes.has(key)) notes.set(key, []);
    notes.get(key).push({ status: flags.length ? `${shift.kind}: ${flags.join(', ')}` : 'По графику', gapText, alert: flags.length > 0 });
  }
  return notes;
}

function plantRows(shifts) {
  const [from, to] = selectedBounds();
  const grouped = new Map();
  for (const shift of shifts) {
    if (shift.start < from || shift.start >= to) continue;
    const period = periodKey(shift.start, $('group').value);
    const key = `${shift.id}|${period}`;
    if (!grouped.has(key)) grouped.set(key, { id: shift.id, name: shift.name, period, plant: 0, visits: [] });
    const row = grouped.get(key);
    row.plant += shift.visits.reduce((sum, visit) => sum + visit.end - visit.start, 0);
    row.visits.push(...shift.visits);
  }
  return [...grouped.values()].sort(bySort);
}

function renderTable(id, headers, rows) {
  $(id).innerHTML = `<thead><tr>${headers.map(h => `<th>${esc(h)}</th>`).join('')}</tr></thead><tbody>${rows.length ? rows.join('') : `<tr><td colspan="${headers.length}" class="empty">Нет данных за выбранный период.</td></tr>`}</tbody>`;
}
function splitDays(start, end, emit) {
  let at = start;
  while (at < end) {
    const d = new Date(at);
    const next = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1).getTime();
    const stop = Math.min(next, end);
    emit(at, stop);
    at = stop;
  }
}
function timelineRows() {
  const [from, to] = selectedBounds(), result = [];
  const add = (item, start, end, type, zone, a, b, kind) => {
    if (end <= from || start >= to) return;
    splitDays(Math.max(start, from), Math.min(end, to), (x, y) => result.push({ id: item.id, name: item.name, period: localDate(x), ts: x, start: x, end: y, ms: y - x, type, zone, kind, a, b }));
  };
  for (const s of analysis.sessions) add(s, s.start, s.end, 'Завод', 'На территории завода', s.entry, s.exit, 'gate');
  for (const s of analysis.segments) add(s, s.start, s.end, s.kind === 'shop' ? 'Цех' : s.kind === 'canteen' ? 'Столовая' : s.kind === 'holding' ? 'Территория' : s.kind === 'unknown' ? 'Несостыковка' : 'Другая зона', s.zone, s.sourceStart, s.sourceEnd, s.kind);
  const daily = buildReport(analysis, $('from-date').value, $('to-date').value, 'day');
  for (const r of daily.rows) if (r.lunchShop) result.push({ id: r.id, name: r.name, period: r.period, ts: new Date(`${r.period}T12:00:00`).getTime(), start: null, end: null, ms: r.lunchShop, type: 'Обед', zone: 'В цехе, расчётно', kind: 'lunch', a: null, b: null });
  const mode = $('sort').value;
  return result.sort((a, b) => {
    const person = collator.compare(a.name, b.name) || collator.compare(a.id, b.id);
    const date = a.period.localeCompare(b.period) * (mode === 'date-desc' ? -1 : 1);
    return (mode === 'name' ? person || date : date || person) || a.ts - b.ts || (a.kind === 'gate' ? -1 : b.kind === 'gate' ? 1 : 0);
  });
}
function shopVisual(row) {
  const entries = report.shops.map((s, i) => ({ s, ms: row.byShop[s] || 0, color: shopColors[i % shopColors.length] })).filter(x => x.ms > 0);
  if (!entries.length) return '<span class="small-muted">Нет подтверждённого времени</span>';
  const bar = entries.map(e => `<span style="flex:${e.ms};background:${e.color}" title="${esc(e.s)} · ${duration(e.ms)}"></span>`).join('');
  const tags = entries.map(e => `<span class="shop-tag" style="--dot:${e.color}">${esc(e.s)} <b>${duration(e.ms)}</b></span>`).join('');
  return `<div class="shop-bar">${bar}</div><div class="shop-tags">${tags}</div>`;
}
function recalculate() {
  const events = dedupeEvents(rawEvents);
  analysis = analyze(events);
  report = buildReport(analysis, $('from-date').value, $('to-date').value, $('group').value);
  const rows = [...report.rows].sort(bySort);
  const shifts = groupSessionsByShift(analysis.sessions);
  const plant = plantRows(shifts);
  const notes = scheduleNotes(shifts);
  const noteFor = r => notes.get(`${r.id}|${r.period}`) || [];
  const plantHeaders = ['Период','Сотрудник','Таб. №','Вход на завод','Выход с завода','На заводе','График'];
  const noteText = r => noteFor(r).map(n => [n.status, n.gapText].filter(Boolean).join('; ')).join('; ');
  const noteHtml = r => noteFor(r).map(n => `<span class="${n.alert ? 'schedule-warning' : 'schedule-ok'}">${esc(n.status)}</span>${n.gapText ? `<span class="schedule-info">${esc(n.gapText)}</span>` : ''}`).join('') || '—';
  const plantData = plant.map(r => [r.period, r.name, r.id, r.visits.map(s => stamp(s.start)).join('\n'), r.visits.map(s => stamp(s.end)).join('\n'), duration(r.plant), noteText(r)]);
  renderTable('plant-table', plantHeaders, plant.map(r => `<tr><td class="datecell">${esc(r.period)}</td><td class="person">${esc(r.name)}</td><td class="small-muted">${esc(r.id)}</td><td><div class="gate-list">${r.visits.map(s => `<span class="gate-pair"><b>↘ ${stamp(s.start)}</b></span>`).join('')}</div></td><td><div class="gate-list">${r.visits.map(s => `<span class="gate-pair"><strong>↗ ${stamp(s.end)}</strong></span>`).join('')}</div></td><td class="num accent">${duration(r.plant)}</td><td class="schedule-cell">${noteHtml(r)}</td></tr>`));
  const shopHeaders = ['Период','Сотрудник','Таб. №','Цеха и время','В цехах','Обед в цехе','Без обеда','Столовая','Вне цехов','Не распределено'];
  const shopData = rows.map(r => [r.period, r.name, r.id, report.shops.filter(s => r.byShop[s] > 0).map(s => `${s}: ${duration(r.byShop[s])}`).join('; '), duration(r.shopGross), duration(r.lunchShop), duration(Math.max(0, r.shopGross - r.lunchShop)), duration(r.canteen), duration(r.holding + r.other), duration(r.unknown)]);
  const shopUiHeaders = ['Период','Сотрудник','Таб. №','Цеха и время','Обед','Без обеда','Вне цехов','Неясно'];
  renderTable('shop-table', shopUiHeaders, rows.map(r => `<tr><td class="datecell">${esc(r.period)}</td><td class="person">${esc(r.name)}</td><td class="small-muted">${esc(r.id)}</td><td class="shop-cell">${shopVisual(r)}</td><td class="num small-muted">${duration(r.lunchShop)}</td><td class="num accent">${duration(Math.max(0, r.shopGross - r.lunchShop))}</td><td class="num">${duration(r.canteen + r.holding + r.other)}</td><td class="num">${duration(r.unknown)}</td></tr>`));
  const timeline = timelineRows();
  const timelineHeaders = ['Период','Сотрудник','Таб. №','Тип','Место / событие','Начало','Конец','Время','Строки Excel'];
  const timelineData = timeline.map(r => [r.period,r.name,r.id,r.type,r.zone,r.start == null ? '—' : stamp(r.start),r.end == null ? '—' : stamp(r.end),duration(r.ms),[r.a,r.b].filter(Boolean).map(source).join('; ')]);
  const groupSize = (index) => {
    let count = 1;
    while (index + count < timeline.length && timeline[index + count].id === timeline[index].id && timeline[index + count].period === timeline[index].period) count++;
    return count;
  };
  renderTable('timeline-table', timelineHeaders, timeline.map((r, index) => {
    const first = index === 0 || timeline[index - 1].id !== r.id || timeline[index - 1].period !== r.period;
    const shared = first ? `<td class="datecell group-cell" rowspan="${groupSize(index)}">${esc(r.period)}</td><td class="person group-cell" rowspan="${groupSize(index)}">${esc(r.name)}</td><td class="small-muted group-cell" rowspan="${groupSize(index)}">${esc(r.id)}</td>` : '';
    return `<tr class="${first ? 'group-start' : ''}">${shared}<td><span class="type-pill ${r.kind === 'shop' ? 'shop' : r.kind === 'gate' ? 'gate' : r.kind === 'unknown' ? 'unknown' : ''}">${esc(r.type)}</span></td><td>${esc(r.zone)}</td><td class="datecell">${r.start == null ? '—' : stamp(r.start)}</td><td class="datecell">${r.end == null ? '—' : stamp(r.end)}</td><td class="num">${duration(r.ms)}</td><td class="source">${[r.a,r.b].filter(Boolean).map(x => esc(source(x))).join('<br>') || '—'}</td></tr>`;
  }));
  $('summary').innerHTML = [[new Set(rows.map(r => r.id)).size,'Сотрудников'],[duration(plant.reduce((n,r) => n+r.plant,0)),'На заводе'],[duration(rows.reduce((n,r) => n+r.shopGross,0)),'В цехах']].map(([v,l]) => `<div class="metric"><b>${esc(v)}</b><span>${esc(l)}</span></div>`).join('');
  exportSheets = [
    {
      name: 'Завод', headers: plantHeaders, rows: plantData,
      columnWidths: [18, 30, 14, 30, 30, 16, 48],
      rowStyles: plant.map(r => noteFor(r).some(n => n.alert) ? excelStyle.redRow : noteFor(r).some(n => n.gapText) ? excelStyle.amberRow : excelStyle.greenRow),
      cellStyles: plant.map(r => [,,,,,, noteFor(r).some(n => n.alert) ? excelStyle.red : noteFor(r).some(n => n.gapText) ? excelStyle.amber : excelStyle.green]),
    },
    {
      name: 'Цеха', headers: ['Период','Сотрудник','Таб. №','Цеха и время',...shopHeaders.slice(4)], rows: shopData,
      columnWidths: [18, 30, 14, 48, 16, 18, 16, 16, 18, 20],
      rowStyles: rows.map((_, index) => index % 2 ? excelStyle.grayRow : excelStyle.plain),
      cellStyles: rows.map(r => [,,, r.shopGross ? excelStyle.green : undefined, undefined, r.lunchShop ? excelStyle.amber : undefined, r.shopGross > r.lunchShop ? excelStyle.green : undefined, r.canteen ? excelStyle.amber : undefined, r.holding + r.other ? excelStyle.blue : undefined, r.unknown ? excelStyle.red : undefined]),
    },
    {
      name: 'Хронология', headers: timelineHeaders, rows: timelineData,
      columnWidths: [18, 30, 14, 18, 35, 28, 28, 14, 48],
      rowStyles: timeline.map(r => r.kind === 'unknown' ? excelStyle.redRow : r.kind === 'shop' ? excelStyle.greenRow : r.kind === 'gate' ? excelStyle.blueRow : r.kind === 'canteen' || r.kind === 'lunch' ? excelStyle.amberRow : excelStyle.grayRow),
      cellStyles: timeline.map(r => [,,, r.kind === 'unknown' ? excelStyle.red : r.kind === 'shop' ? excelStyle.green : r.kind === 'gate' ? excelStyle.blue : r.kind === 'canteen' || r.kind === 'lunch' ? excelStyle.amber : undefined]),
    },
  ];
  notify('');
}

$('files').addEventListener('change', async event => {
  const selected=[...event.target.files]; if (!selected.length) return;
  notify('Читаю файлы…'); let added=0,rejected=0;
  for (const file of selected) {
    try { const result=parseEvents(await readXlsx(file),file.name); rawEvents.push(...result.events); files.push(`${file.name} · ${result.events.length} отметок`); added+=result.events.length; rejected+=result.rejected.length; }
    catch(error) { files.push(`${file.name} · ошибка: ${error.message}`); }
  }
  $('file-list').textContent=files.join('  /  '); $('files').value='';
  if (!rawEvents.length) { notify('Нет событий для отчёта.'); return; }
  const sorted=dedupeEvents(rawEvents);
  if ($('workspace').hidden) { $('from-date').value=localDate(sorted[0].ts); $('to-date').value=localDate(sorted.at(-1).ts); }
  $('workspace').hidden=false; $('clear-data').hidden=false;
  try { recalculate(); notify(`Загружено ${added} отметок${rejected ? `; некорректных строк: ${rejected}` : ''}.`); } catch(error) { notify(error.message); }
});
$('calculate').addEventListener('click', () => { try { recalculate(); } catch(error) { notify(error.message); } });
$('sort').addEventListener('change', () => { if (analysis) recalculate(); });
$('group').addEventListener('change', () => { if (analysis) recalculate(); });
$('export-all').addEventListener('click', async () => { try { recalculate(); await downloadWorkbook(exportSheets,`PERCo_${$('from-date').value}_${$('to-date').value}.xlsx`); } catch(error) { notify(`Не удалось создать Excel: ${error.message}`); } });
document.querySelectorAll('.tab').forEach(tab => tab.addEventListener('click', () => {
  document.querySelectorAll('.tab').forEach(x => { const active=x===tab; x.classList.toggle('active',active); x.setAttribute('aria-selected',String(active)); });
  document.querySelectorAll('[data-panel]').forEach(panel => { panel.hidden=panel.dataset.panel!==tab.dataset.tab; });
}));
$('clear-data').addEventListener('click', () => {
  if (!confirm('Удалить загруженные данные?')) return;
  rawEvents=[]; files=[]; analysis=null; report=null; exportSheets=[];
  $('workspace').hidden=true; $('clear-data').hidden=true; $('file-list').textContent='Файл не выбран'; notify('Данные удалены.');
});
