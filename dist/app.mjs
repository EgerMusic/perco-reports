import { parseEvents, duration, localDate } from './core.mjs';
import { readXlsx } from './xlsx.mjs';
import { parseSchedule, parsePresence, buildShiftReport, reportCoverage } from './plan.mjs';
import { downloadWorkbook, excelStyle as S } from './export.mjs';

const $ = id => document.getElementById(id);
const collator = new Intl.Collator('ru', { sensitivity: 'base', numeric: true });
let rawEvents = [], eventFiles = [], schedule = null, presence = null, exportSheets = [], defaultEnd = '';
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
const headers = ['Дата смены','Сотрудник','Таб. №','Норма смены','Первый вход','Последний выход','От входа до выхода','Статус / комментарий'];
const sourceHeaders = ['Основание учёта','Месячный график','Источник нормы','Источники входа и выхода'];
const reportTabs = [
  { id: 'plant-table', name: 'Подробные проходы', factory: false },
  { id: 'presence-table', name: 'Сводка PERCo', factory: true },
];

function makeReportSheet(tab, rows) {
  renderTable(tab.id, headers, rows.map(row => {
    const details = comments(row).map(text => `<span class="schedule-info">${esc(text)}</span>`).join('');
    return `<tr class="result-${row.level}">
      <td class="datecell">${esc(row.date)}</td>
      <td class="person">${esc(row.name)}<span class="sub">${esc(row.department)}</span></td>
      <td class="small-muted">${esc(row.id)}</td>
      <td class="plan-label" title="${esc(row.plan.normSource || row.planSource)}">${esc(row.plan.label)}<span class="sub">${esc(row.factSource)}</span></td>
      <td class="datecell" title="${esc(source(row.first))}"><b>${stamp(row.first?.ts)}</b></td>
      <td class="datecell" title="${esc(source(row.last))}">${stamp(row.last?.ts)}</td>
      <td class="num accent">${fmt(row.span)}</td>
      <td class="schedule-cell"><span class="status-${row.level}">${esc(row.status)}</span>${details}</td>
    </tr>`;
  }));

  return {
    name: tab.name,
    headers: [...headers, ...sourceHeaders],
    rows: rows.map(row => [
      row.date, row.name, row.id, row.plan.label, stamp(row.first?.ts), stamp(row.last?.ts), fmt(row.span),
      [row.status, ...comments(row)].join('\n'), row.factSource, row.planSource, row.plan.normSource || '',
      [...new Set([source(row.first),source(row.last)].filter(Boolean))].join('\n'),
    ]),
    columnWidths: [15, 30, 12, 27, 24, 24, 23, 74, 26, 43, 43, 49],
    rowStyles: rows.map(row => styleFor(row.level)),
    cellStyles: rows.map(row => ({ 7: styleStrong(row.level) })),
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
  if (schedule) {
    const monthFrom = `${schedule.year}-${String(schedule.month).padStart(2,'0')}-01`;
    const monthTo = localDate(new Date(schedule.year,schedule.month,0,12).getTime());
    if (!range || from > monthTo || to < monthFrom) {
      from = monthFrom;
      to = monthTo;
    } else {
      if (from < monthFrom) from = monthFrom;
      if (to > monthTo) to = monthTo;
    }
  }
  if (resetDates || !$('from-date').value) {
    $('from-date').value = from;
    $('to-date').value = to;
  } else if ($('to-date').value === defaultEnd) {
    $('to-date').value = to;
  }
  defaultEnd = to;
  recalculate();
}
$('files').addEventListener('change',async event=>{
  const selected=[...event.target.files]; if (!selected.length)return;
  notify('Читаю проходы…'); const errors=[]; let rejected=0;
  for (const file of selected) try {
    const result=parseEvents(await readXlsx(file),file.name); rawEvents.push(...result.events); rejected+=result.rejected.length;
    eventFiles.push(`${file.name} · ${result.events.length} отметок`);
  } catch(error){ errors.push(`${file.name}: ${error.message}`); }
  $('file-list').textContent=eventFiles.join(' / ')||'Файл не выбран'; $('files').value='';
  try { refresh(); notify([...errors,rejected?`Некорректных строк: ${rejected}`:''].filter(Boolean).join(' ')); } catch(error){notify(error.message);}
});
for (const [inputId,type,label] of [['schedule-file','schedule','schedule-label'],['presence-file','presence','presence-label']]) $(inputId).addEventListener('change',async event=>{
  const file=event.target.files[0];if(!file)return;notify('Читаю файл…');
  try {
    const rows=await readXlsx(file,type);
    if(type==='schedule')schedule=parseSchedule(rows,file.name); else presence=parsePresence(rows,file.name);
    $(label).textContent=`${file.name} · ${type==='schedule'?schedule.people.length:presence.people.length} сотрудников`;
    refresh(type==='schedule'); notify(type==='schedule'&&schedule.warnings.length?`В графике есть повторяющиеся табельные номера: ${schedule.warnings.length}. Они помечены в строках отчёта.`:'');
  }catch(error){notify(error.message);}finally{event.target.value='';}
});
for (const id of ['calculate','sort']) $(id).addEventListener(id==='calculate'?'click':'change',()=>{try{recalculate();notify('');}catch(error){notify(error.message);}});
$('export-all').addEventListener('click',async()=>{try{recalculate();await downloadWorkbook(exportSheets,`PERCo_${$('from-date').value}_${$('to-date').value}.xlsx`);}catch(error){notify(`Не удалось создать Excel: ${error.message}`);}});
document.querySelectorAll('.tab').forEach(tab=>tab.addEventListener('click',()=>{
  document.querySelectorAll('.tab').forEach(x=>{const active=x===tab;x.classList.toggle('active',active);x.setAttribute('aria-selected',String(active));});
  document.querySelectorAll('[data-panel]').forEach(panel=>{panel.hidden=panel.dataset.panel!==tab.dataset.tab;});
}));
$('clear-data').addEventListener('click',()=>{
  if(!confirm('Удалить загруженные данные?'))return;
  rawEvents=[];eventFiles=[];schedule=null;presence=null;exportSheets=[];defaultEnd='';
  $('workspace').hidden=true;$('clear-data').hidden=true;
  for(const id of ['file-list','schedule-label','presence-label'])$(id).textContent='Файл не выбран';
  $('from-date').value='';$('to-date').value='';$('from-date').max='';$('to-date').max='';notify('Данные удалены.');
});
