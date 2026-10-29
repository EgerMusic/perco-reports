import { parseEvents, dedupeEvents, analyze, buildReport, duration, localDate, periodKey } from './core.mjs';
import { readXlsx } from './xlsx.mjs';
import { downloadWorkbook } from './export.mjs';

const $ = id => document.getElementById(id);
const storageKey = 'perco-time-corrections-v1';
const collator = new Intl.Collator('ru', { sensitivity: 'base', numeric: true });
const shopColors = ['#eb8135','#2c8874','#4d7ea0','#9b755e','#836eac','#60926b','#ba786a','#7899a8'];
let rawEvents = [], files = [], corrections = {}, excluded = {}, analysis = null, report = null, exportSheets = [];
try { ({ corrections = {}, excluded = {} } = JSON.parse(localStorage.getItem(storageKey) || '{}')); } catch { corrections = {}; excluded = {}; }
const save = () => localStorage.setItem(storageKey, JSON.stringify({ corrections, excluded }));
const esc = value => String(value ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);
const stamp = ms => new Date(ms).toLocaleString('ru-RU', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
const localInput = ms => { const d = new Date(ms); return `${localDate(ms)}T${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`; };
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
    splitDays(Math.max(start, from), Math.min(end, to), (x, y) => result.push({ id: item.id, name: item.name, period: periodKey(x, $('group').value), ts: x, start: x, end: y, ms: y - x, type, zone, kind, a, b }));
  };
  for (const s of analysis.sessions) add(s, s.start, s.end, 'Завод', 'На территории завода', s.entry, s.exit, 'gate');
  for (const s of analysis.segments) add(s, s.start, s.end, s.kind === 'shop' ? 'Цех' : s.kind === 'canteen' ? 'Столовая' : s.kind === 'holding' ? 'Территория' : s.kind === 'unknown' ? 'Несостыковка' : 'Другая зона', s.zone, s.sourceStart, s.sourceEnd, s.kind);
  for (const r of report.rows) if (r.lunchShop) result.push({ id: r.id, name: r.name, period: r.period, ts: new Date(`${periodDate(r.period)}T00:00:00`).getTime(), start: null, end: null, ms: r.lunchShop, type: 'Обед', zone: 'В цехе, расчётно', kind: 'lunch', a: null, b: null });
  return result.sort(bySort);
}
function shopVisual(row) {
  const entries = report.shops.map((s, i) => ({ s, ms: row.byShop[s] || 0, color: shopColors[i % shopColors.length] })).filter(x => x.ms > 0);
  if (!entries.length) return '<span class="small-muted">Нет подтверждённого времени</span>';
  const bar = entries.map(e => `<span style="flex:${e.ms};background:${e.color}" title="${esc(e.s)} · ${duration(e.ms)}"></span>`).join('');
  const tags = entries.map(e => `<span class="shop-tag" style="--dot:${e.color}">${esc(e.s)} <b>${duration(e.ms)}</b></span>`).join('');
  return `<div class="shop-bar">${bar}</div><div class="shop-tags">${tags}</div>`;
}
function inRange(issue) {
  const [from, to] = selectedBounds();
  return (issue.a.ts >= from && issue.a.ts < to) || (issue.b && issue.b.ts >= from && issue.b.ts < to) || (issue.b && issue.a.ts < from && issue.b.ts >= to);
}

function recalculate() {
  const events = dedupeEvents(rawEvents).filter(e => !excluded[e.key]);
  analysis = analyze(events, corrections);
  report = buildReport(analysis, $('from-date').value, $('to-date').value, $('group').value);
  const rows = [...report.rows].sort(bySort);
  const issues = analysis.issues.filter(inRange);
  const plantHeaders = ['Период','Сотрудник','Таб. №','Вход на завод','Выход с завода','На заводе'];
  const plantData = rows.map(r => [r.period, r.name, r.id, r.visits.map(s => stamp(s.start)).join('\n'), r.visits.map(s => stamp(s.end)).join('\n'), duration(r.plant)]);
  renderTable('plant-table', plantHeaders, rows.map(r => `<tr><td class="datecell">${esc(r.period)}</td><td class="person">${esc(r.name)}</td><td class="small-muted">${esc(r.id)}</td><td><div class="gate-list">${r.visits.map(s => `<span class="gate-pair"><b>↘ ${stamp(s.start)}</b></span>`).join('')}</div></td><td><div class="gate-list">${r.visits.map(s => `<span class="gate-pair"><strong>↗ ${stamp(s.end)}</strong></span>`).join('')}</div></td><td class="num accent">${duration(r.plant)}</td></tr>`));
  const shopHeaders = ['Период','Сотрудник','Таб. №','Цеха и время','В цехах','Обед в цехе','Без обеда','Столовая','Вне цехов','Не распределено'];
  const shopData = rows.map(r => [r.period, r.name, r.id, report.shops.filter(s => r.byShop[s] > 0).map(s => `${s}: ${duration(r.byShop[s])}`).join('; '), duration(r.shopGross), duration(r.lunchShop), duration(Math.max(0, r.shopGross - r.lunchShop)), duration(r.canteen), duration(r.holding + r.other), duration(r.unknown)]);
  const shopUiHeaders = ['Период','Сотрудник','Таб. №','Цеха и время','Обед','Без обеда','Вне цехов','Неясно'];
  renderTable('shop-table', shopUiHeaders, rows.map(r => `<tr><td class="datecell">${esc(r.period)}</td><td class="person">${esc(r.name)}</td><td class="small-muted">${esc(r.id)}</td><td class="shop-cell">${shopVisual(r)}</td><td class="num small-muted">${duration(r.lunchShop)}</td><td class="num accent">${duration(Math.max(0, r.shopGross - r.lunchShop))}</td><td class="num">${duration(r.canteen + r.holding + r.other)}</td><td class="num">${duration(r.unknown)}</td></tr>`));
  const timeline = timelineRows();
  const timelineHeaders = ['Период','Сотрудник','Таб. №','Тип','Место / событие','Начало','Конец','Время','Строки Excel'];
  const timelineData = timeline.map(r => [r.period,r.name,r.id,r.type,r.zone,r.start == null ? '—' : stamp(r.start),r.end == null ? '—' : stamp(r.end),duration(r.ms),[r.a,r.b].filter(Boolean).map(source).join('; ')]);
  renderTable('timeline-table', timelineHeaders, timeline.map(r => `<tr><td class="datecell">${esc(r.period)}</td><td class="person">${esc(r.name)}</td><td class="small-muted">${esc(r.id)}</td><td><span class="type-pill ${r.kind === 'shop' ? 'shop' : r.kind === 'gate' ? 'gate' : r.kind === 'unknown' ? 'unknown' : ''}">${esc(r.type)}</span></td><td>${esc(r.zone)}</td><td class="datecell">${r.start == null ? '—' : stamp(r.start)}</td><td class="datecell">${r.end == null ? '—' : stamp(r.end)}</td><td class="num">${duration(r.ms)}</td><td class="source">${[r.a,r.b].filter(Boolean).map(x => esc(source(x))).join('<br>') || '—'}</td></tr>`));
  $('summary').innerHTML = [[new Set(rows.map(r => r.id)).size,'Сотрудников'],[duration(rows.reduce((n,r) => n+r.plant,0)),'На заводе'],[duration(rows.reduce((n,r) => n+r.shopGross,0)),'В цехах'],[issues.length,'Несостыковок']].map(([v,l]) => `<div class="metric"><b>${esc(v)}</b><span>${esc(l)}</span></div>`).join('');
  exportSheets = [
    { name: 'Завод', headers: plantHeaders, rows: plantData },
    { name: 'Цеха', headers: ['Период','Сотрудник','Таб. №','Цеха и время',...shopHeaders.slice(4)], rows: shopData },
    { name: 'Хронология', headers: timelineHeaders, rows: timelineData },
    { name: 'Несостыковки', headers: ['Сотрудник','Таб. №','Тип','Отметка 1','Строка Excel 1','Отметка 2','Строка Excel 2'], rows: issues.map(i => [i.name,i.id,issueTitle(i),stamp(i.a.ts),source(i.a),i.b ? stamp(i.b.ts) : '',i.b ? source(i.b) : '']) },
  ];
  renderIssues(issues);
  notify('');
}

function issueTitle(issue) { return ({ 'zone-gap':'Не отмечен переход между зонами','repeat-entry':'Повторный вход без выхода','orphan-exit':'Выход без входа в данных','open-entry':'Вход без последующего выхода' })[issue.type] || issue.type; }
function renderIssues(visible) {
  $('issue-count').textContent = visible.length;
  if (!visible.length) $('issues').innerHTML = '<div class="empty">Несостыковок за выбранный период нет.</div>';
  else $('issues').innerHTML = visible.map(issue => {
    const index = analysis.issues.indexOf(issue);
    const item = e => `<div>${stamp(e.ts)} · ${esc(e.from)} → ${esc(e.to)} · <button class="source-button" type="button" data-copy="${esc(source(e))}">${esc(source(e))}</button></div>`;
    const fixable = issue.type === 'zone-gap' && issue.a.to !== 'Неконтролируемая территория' && issue.b.from !== 'Неконтролируемая территория';
    const form = fixable ? `<div class="issue-form"><label>Время перехода<input type="datetime-local" data-time="${index}" min="${localInput(issue.a.ts)}" max="${localInput(issue.b.ts)}"></label><label>Основание<input type="text" data-reason="${index}" placeholder="Со слов мастера"></label><button class="primary" type="button" data-fix="${index}">Применить</button></div>` : `<div class="issue-form"><label>Причина исключения<input type="text" data-reason="${index}" placeholder="Ошибка прохода"></label><button class="secondary" type="button" data-exclude="${index}" data-which="a">Исключить отметку${issue.type === 'repeat-entry' ? ' 1' : ''}</button>${issue.type === 'repeat-entry' ? `<button class="secondary" type="button" data-exclude="${index}" data-which="b">Исключить отметку 2</button>` : ''}</div>`;
    return `<article class="issue"><div class="issue-head"><span class="chip">${esc(issue.id)}</span><strong>${esc(issue.name || 'Сотрудник')}</strong><span>${esc(issueTitle(issue))}</span></div><div class="details">${item(issue.a)}${issue.b ? item(issue.b) : ''}</div>${form}</article>`;
  }).join('');
  renderChanges();
}
function renderChanges() {
  const changes = [...Object.entries(corrections).map(([key,value]) => ({key,type:'correction',label:`Переход ${key.split('|')[0]} · ${stamp(Number(value.cut ?? value))}`,reason:value.reason || ''})),...Object.entries(excluded).map(([key,value]) => ({key,type:'excluded',label:`Исключена отметка ${key.split('|')[0]}`,reason:value.reason || ''}))];
  $('changes').innerHTML = changes.length ? `<h3>Ручные исправления</h3>${changes.map((c,i) => `<div class="issue"><div class="issue-head"><strong>${esc(c.label)}</strong><span>${esc(c.reason)}</span></div><button class="text-button" type="button" data-undo="${i}">Отменить</button></div>`).join('')}` : '';
  $('changes').querySelectorAll('[data-undo]').forEach(button => button.addEventListener('click', () => { const c=changes[Number(button.dataset.undo)]; if(c.type==='correction') delete corrections[c.key]; else delete excluded[c.key]; save(); recalculate(); }));
}
$('issues').addEventListener('click', async event => {
  const copy = event.target.closest('[data-copy]');
  if (copy) { try { await navigator.clipboard.writeText(copy.dataset.copy); notify(`Скопировано: ${copy.dataset.copy}`); } catch { notify(copy.dataset.copy); } return; }
  const fix = event.target.closest('[data-fix]');
  if (fix) {
    const i=Number(fix.dataset.fix), issue=analysis.issues[i], cut=new Date($('issues').querySelector(`[data-time="${i}"]`).value).getTime();
    if (!Number.isFinite(cut) || cut <= issue.a.ts || cut >= issue.b.ts) { notify('Укажи время строго между двумя отметками.'); return; }
    const reason=$('issues').querySelector(`[data-reason="${i}"]`).value.trim();
    if (!reason) { notify('Укажи основание исправления.'); return; }
    corrections[issue.key]={cut,reason,at:new Date().toISOString()}; save(); recalculate(); return;
  }
  const exclude=event.target.closest('[data-exclude]');
  if (exclude) {
    const i=Number(exclude.dataset.exclude), issue=analysis.issues[i], reason=$('issues').querySelector(`[data-reason="${i}"]`).value.trim();
    if (!reason) { notify('Укажи причину исключения.'); return; }
    const e=issue[exclude.dataset.which]; excluded[e.key]={reason,at:new Date().toISOString()}; save(); recalculate();
  }
});
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
  if (!confirm('Удалить загруженные данные и ручные исправления?')) return;
  rawEvents=[]; files=[]; corrections={}; excluded={}; analysis=null; report=null; exportSheets=[]; localStorage.removeItem(storageKey);
  $('workspace').hidden=true; $('clear-data').hidden=true; $('file-list').textContent='Файл не выбран'; notify('Данные удалены.');
});
