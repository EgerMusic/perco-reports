import { parseEvents, duration, localDate } from './core.mjs';
import { readWorkbook } from './xlsx.mjs';
import { parseSchedule, mergeSchedules, parsePresence, buildShiftReport, reportCoverage } from './plan.mjs';
import { downloadWorkbook, excelStyle as S } from './export.mjs';

const $ = id => document.getElementById(id);
const collator = new Intl.Collator('ru', { sensitivity: 'base', numeric: true });
let rawEvents = [], eventFiles = [], scheduleFiles = [], schedule = null, presence = null, exportSheets = [];
let defaultStart = '', defaultEnd = '';
const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'})[c]);
const stamp = ts => ts == null ? '—' : new Date(ts).toLocaleString('ru-RU', {day:'2-digit',month:'2-digit',year:'numeric',hour:'2-digit',minute:'2-digit'});
const fmt = ms => ms == null ? '—' : duration(ms);
const source = e => e ? `${e.fileName}${e.sheetName ? ` · лист ${e.sheetName}` : ''} · строка ${e.rowNumber}` : '';
const notify = text => { $('notice').textContent = text; };
// Remove copies kept by the previous release. This version only uses page memory.
try {
  const scope = location.pathname.replace(/\/index\.html$/i, '').replace(/\/+$/, '') || '/';
  globalThis.indexedDB?.deleteDatabase(`perco-workspace-v1:${scope}`);
} catch { /* Processing files also works when browser storage is unavailable. */ }
const bySort = (a, b) => {
  const person = collator.compare(a.name, b.name) || collator.compare(a.employeeKey, b.employeeKey);
  const date = a.date.localeCompare(b.date) * ($('sort').value === 'date-desc' ? -1 : 1);
  return $('sort').value === 'name' ? person || date : date || person;
};
const styleFor = level => level === 'red' ? S.redRow : level === 'amber' ? S.amberRow : level === 'green' ? S.greenRow : S.grayRow;
const styleStrong = level => level === 'red' ? S.red : level === 'amber' ? S.amber : level === 'green' ? S.green : S.plain;

function renderTable(id, headers, rows) {
  $(id).innerHTML = `<thead><tr>${headers.map(h => `<th>${esc(h)}</th>`).join('')}</tr></thead><tbody>${rows.length ? rows.join('') : `<tr><td colspan="${headers.length}" class="empty">Нет данных за выбранный период.</td></tr>`}</tbody>`;
}
function gapText(row) {
  if (!row.gaps.length) return '';
  const sum = row.gaps.reduce((n, v) => n + v, 0);
  return row.gaps.length === 1 ? `Выходил с завода на ${duration(sum)}` : `Выходил с завода ${row.gaps.length} раз, всего ${duration(sum)}`;
}
const comments = r => [...r.comments, gapText(r)].filter(Boolean);
const headers = ['Дата смены','Сотрудник','Таб. №','Норма смены','Первый вход','Последний выход','Часы без обеда','Статус / комментарий'];
const sourceHeaders = ['Основание учёта','Месячный график','Источник нормы','Источники входа и выхода'];
const reportTabs = [
  { id: 'plant-table', name: 'Подробные проходы', factory: false },
  { id: 'presence-table', name: 'Сводка PERCo', factory: true },
];

function makeReportSheet(tab, rows) {
  const totals = new Map();
  rows.forEach((row, index) => {
    const count = totals.get(row.employeeKey)?.count || 0;
    totals.set(row.employeeKey, { last:index, count:count + Number(row.worked > 0) });
  });
  const lines = rows.flatMap((row, index) => {
    const { last, count } = totals.get(row.employeeKey);
    return index === last ? [{ row }, { row, count }] : [{ row }];
  });
  const totalText = count => `Отработано смен — ${count}`;
  renderTable(tab.id, headers, lines.map(({ row, count }) => {
    if (count !== undefined) {
      return `<tr class="employee-total"><td colspan="6">Итого за период · ${esc(row.name)} · № ${esc(row.id)}</td><td colspan="2">${totalText(count)}</td></tr>`;
    }
    const details = comments(row).map(text => `<span class="schedule-info">${esc(text)}</span>`).join('');
    return `<tr class="result-${row.level}">
      <td class="datecell">${esc(row.date)}</td>
      <td class="person">${esc(row.name)}<span class="sub">${esc(row.department)}</span></td>
      <td class="small-muted">${esc(row.id)}</td>
      <td class="plan-label" title="${esc(row.plan.normSource || row.planSource)}">${esc(row.plan.label)}<span class="sub">${esc(row.factSource)}</span></td>
      <td class="datecell" title="${esc(source(row.first))}"><b>${stamp(row.first?.ts)}</b></td>
      <td class="datecell" title="${esc(source(row.last))}">${stamp(row.last?.ts)}</td>
      <td class="num accent">${fmt(row.worked)}</td>
      <td class="schedule-cell"><span class="status-${row.level}">${esc(row.status)}</span>${details}</td>
    </tr>`;
  }));

  return {
    name: tab.name,
    headers: [...headers, ...sourceHeaders],
    rows: lines.map(({ row, count }) => {
      if (count !== undefined) return ['Итого за период', row.name, row.id, '', '', '', '', totalText(count), '', '', '', ''];
      return [
        row.date, row.name, row.id, row.plan.label, stamp(row.first?.ts), stamp(row.last?.ts), fmt(row.worked),
        [row.status, ...comments(row)].join('\n'), row.factSource, row.planSource, row.plan.normSource || '',
        [...new Set([source(row.first),source(row.last)].filter(Boolean))].join('\n'),
      ];
    }),
    columnWidths: [18, 30, 12, 27, 24, 24, 23, 74, 26, 43, 43, 49],
    rowStyles: lines.map(({ row, count }) => count === undefined ? styleFor(row.level) : S.blueRow),
    cellStyles: lines.map(({ row, count }) => ({ 7: count === undefined ? styleStrong(row.level) : S.blue })),
  };
}

function recalculate() {
  const lastDate = $('to-date').max;
  if (lastDate && $('to-date').value > lastDate) $('to-date').value = lastDate;
  const report = buildShiftReport(rawEvents, schedule, presence, $('from-date').value, $('to-date').value);
  const rows = [...report.rows].sort(bySort);
  const summaryRows = rows.filter(row => row.accounting === 'factory');
  $('export-all').disabled = !rows.length;
  $('summary').innerHTML = [
    [new Set(rows.map(row => row.employeeKey)).size, 'Сотрудников'],
    [rows.length, 'Строк в отчёте'],
    [rows.filter(row => row.level === 'red').length, 'Смен с отклонениями'],
    [rows.filter(row => row.needsReview).length, 'Строк для ручного рассмотрения'],
  ].map(([value, label]) => `<div class="metric"><b>${value}</b><span>${label}</span></div>`).join('');
  exportSheets = reportTabs.map(tab => makeReportSheet(tab, tab.factory ? summaryRows : rows.filter(row => row.accounting !== 'factory')));
}

function refresh(resetDates = false) {
  const range = reportCoverage(rawEvents, presence);
  $('workspace').hidden = !range && !schedule;
  $('clear-data').hidden = !rawEvents.length && !schedule && !presence;
  if ($('workspace').hidden) return;
  const lastDate = range ? localDate(range.end) : '';
  $('from-date').max = lastDate;
  $('to-date').max = lastDate;
  let from = range ? localDate(range.start) : '', to = lastDate;
  if (schedule && !range) {
    from = schedule.fromDate;
    to = schedule.toDate;
  } else if (schedule && from <= schedule.toDate && to >= schedule.fromDate) {
    if (from < schedule.fromDate) from = schedule.fromDate;
    if (to > schedule.toDate) to = schedule.toDate;
  }
  if (resetDates || !$('from-date').value || lastDate && $('from-date').value > lastDate) {
    $('from-date').value = from;
    $('to-date').value = to;
  } else {
    if ($('from-date').value === defaultStart) $('from-date').value = from;
    if ($('to-date').value === defaultEnd) $('to-date').value = to;
  }
  defaultStart = from;
  defaultEnd = to;
  recalculate();
}
$('files').addEventListener('change',async event=>{
  const selected=[...event.target.files]; if (!selected.length)return;
  notify('Читаю проходы…'); const errors=[]; let rejected=0;
  for (const file of selected) try {
    const result=parseEvents(await readWorkbook(file),file.name); rawEvents.push(...result.events); rejected+=result.rejected.length;
    eventFiles.push(`${file.name} · ${result.events.length} отметок`);
  } catch(error){ errors.push(`${file.name}: ${error.message}`); }
  $('file-list').textContent=eventFiles.join(' / ')||'Файл не выбран'; $('files').value='';
  try { refresh(); notify([...errors,rejected?`Некорректных строк: ${rejected}`:''].filter(Boolean).join(' ')); } catch(error){notify(error.message);}
});
$('schedule-file').addEventListener('change',async event=>{
  const input = event.target, files = [...input.files];
  if (!files.length) return;
  const previous = scheduleFiles, added = [], errors = [];
  input.disabled = true;
  notify('Читаю месячные графики…');
  try {
    for (const file of files) {
      try { added.push(parseSchedule(await readWorkbook(file, 'schedule'), file.name)); }
      catch(error) { errors.push(`${file.name}: ${error.message}`); }
    }
    if (previous !== scheduleFiles) return;
    if (added.length) {
      const updated = [...previous, ...added];
      schedule = mergeSchedules(updated);
      scheduleFiles = updated;
      const filenames = updated.map(file => file.fileName).join(' / ');
      $('schedule-label').textContent = filenames;
      $('schedule-label').title = filenames;
      refresh(true);
    }
    if (schedule?.warnings.length) errors.push(`Замечаний в графиках — ${schedule.warnings.length}. Подробности в строках отчёта.`);
    notify(errors.join(' '));
  } catch(error) { notify(error.message); }
  finally { input.value = ''; input.disabled = false; }
});
$('presence-file').addEventListener('change',async event=>{
  const file=event.target.files[0];if(!file)return;notify('Читаю файл…');
  try {
    presence=parsePresence(await readWorkbook(file,'presence'),file.name);
    $('presence-label').textContent=`${file.name} · ${presence.people.length} сотрудников`;
    refresh();notify('');
  } catch(error) {notify(error.message);} finally {event.target.value='';}
});
for (const id of ['calculate','sort']) $(id).addEventListener(id==='calculate'?'click':'change',()=>{try{recalculate();notify('');}catch(error){notify(error.message);}});
$('export-all').addEventListener('click',async()=>{try{recalculate();await downloadWorkbook(exportSheets,`PERCo_${$('from-date').value}_${$('to-date').value}.xlsx`);}catch(error){notify(`Не удалось создать Excel: ${error.message}`);}});
document.querySelectorAll('.tab').forEach(tab=>tab.addEventListener('click',()=>{
  document.querySelectorAll('.tab').forEach(x=>{const active=x===tab;x.classList.toggle('active',active);x.setAttribute('aria-selected',String(active));});
  document.querySelectorAll('[data-panel]').forEach(panel=>{panel.hidden=panel.dataset.panel!==tab.dataset.tab;});
}));
$('clear-data').addEventListener('click',()=>{
  if(!confirm('Удалить загруженные данные?'))return;
  rawEvents=[];eventFiles=[];scheduleFiles=[];schedule=null;presence=null;exportSheets=[];defaultStart='';defaultEnd='';
  $('workspace').hidden=true;$('clear-data').hidden=true;
  for(const id of ['file-list','presence-label'])$(id).textContent='Файл не выбран';
  $('schedule-label').textContent='Файлы не выбраны';
  $('schedule-label').removeAttribute('title');
  $('from-date').value='';$('to-date').value='';$('from-date').max='';$('to-date').max='';notify('Данные удалены.');
});
