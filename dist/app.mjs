import { parseEvents, dedupeEvents, duration, localDate } from './core.mjs';
import { readXlsx } from './xlsx.mjs';
import { parseSchedule, parsePresence, buildShiftReport, aggregateShops } from './plan.mjs';
import { downloadWorkbook, excelStyle as S } from './export.mjs';

const $ = id => document.getElementById(id);
const collator = new Intl.Collator('ru', { sensitivity: 'base', numeric: true });
const shopColors = ['#eb8135','#2c8874','#4d7ea0','#9b755e','#836eac','#60926b','#ba786a','#7899a8'];
let rawEvents = [], eventFiles = [], schedule = null, presence = null, shiftReport = null, exportSheets = [];
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
const dateValue = r => r.date || r.period.replace('Неделя с ', '');
const bySort = (a, b) => {
  const person = collator.compare(a.name, b.name) || collator.compare(a.employeeKey, b.employeeKey);
  const date = dateValue(a).localeCompare(dateValue(b)) * ($('sort').value === 'date-desc' ? -1 : 1);
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
function shopVisual(row) {
  const entries = shiftReport.shops.map((s, i) => ({s, ms: row.byShop[s] || 0, color: shopColors[i % shopColors.length]})).filter(e => e.ms > 0);
  if (!entries.length) return '<span class="small-muted">Нет подтверждённого времени</span>';
  return `<div class="shop-bar">${entries.map(e => `<span style="flex:${e.ms};background:${e.color}" title="${esc(e.s)} · ${duration(e.ms)}"></span>`).join('')}</div><div class="shop-tags">${entries.map(e => `<span class="shop-tag" style="--dot:${e.color}">${esc(e.s)} <b>${duration(e.ms)}</b></span>`).join('')}</div>`;
}
function timelineRows(rows) {
  const result = [];
  for (const r of rows) {
    for (const s of r.segments) result.push({ ...r, start:s.start, end:s.end, ms:s.end-s.start, zone:s.zone, kind:s.kind, sources:[source(s.a),source(s.b)].join('; ') });
    for (const v of r.visits) result.push({ ...r, start:v.start, end:v.end, ms:v.end-v.start, zone:'На территории завода', kind:'gate', sources:'' });
    if (r.lunchShop) result.push({ ...r, start:null, end:null, ms:r.lunchShop, zone:'Обед в цехе, расчётно', kind:'lunch', sources:'' });
  }
  return result.sort((a, b) => bySort(a,b) || (a.start ?? Infinity) - (b.start ?? Infinity) || (a.kind === 'gate' ? -1 : 1));
}

function recalculate() {
  const lastDate = $('to-date').max;
  if (lastDate && $('to-date').value > lastDate) $('to-date').value = lastDate;
  shiftReport = buildShiftReport(rawEvents, schedule, presence, $('from-date').value, $('to-date').value);
  const rows = [...shiftReport.rows].sort(bySort), shops = aggregateShops(shiftReport, $('group').value), shopRows = shops.rows.sort(bySort);
  $('export-all').disabled = !rows.length;
  const headers = ['Дата смены','Сотрудник','Таб. №','План','Первый вход в цех','Последний выход из цеха','От входа до выхода','В цехах','Обед в цехе','Без обеда','График / комментарий','Источник плана','Источники входа и выхода'];
  const data = rows.map(r => [r.date,r.name,r.id,r.plan.label,stamp(r.first?.ts),stamp(r.last?.ts),fmt(r.span),fmt(r.observed?r.shopGross:null),fmt(r.observed?r.lunchShop:null),fmt(r.observed?Math.max(0,r.shopGross-r.lunchShop):null),[r.status,...comments(r)].join('\n'),r.planSource,[source(r.first),source(r.last)].filter(Boolean).join('\n')]);
  const uiHeaders = headers.filter((_,i)=>![7,8,9,11,12].includes(i));
  renderTable('plant-table',uiHeaders,rows.map(r => `<tr class="result-${r.level}"><td class="datecell">${esc(r.date)}</td><td class="person">${esc(r.name)}<span class="sub">${esc(r.department)}</span></td><td class="small-muted">${esc(r.id)}</td><td class="plan-label" title="${esc(r.planSource)}">${esc(r.plan.label)}</td><td class="datecell" title="${esc(source(r.first))}"><b>${stamp(r.first?.ts)}</b></td><td class="datecell" title="${esc(source(r.last))}">${stamp(r.last?.ts)}</td><td class="num accent">${fmt(r.span)}</td><td class="schedule-cell"><span class="status-${r.level}">${esc(r.status)}</span>${comments(r).map(text=>`<span class="schedule-info">${esc(text)}</span>`).join('')}</td></tr>`));
  const shopHeaders = ['Период','Сотрудник','Таб. №','Цеха и время','В цехах','Обед в цехе','Без обеда','Столовая','Вне цехов','Неясно'];
  const shopData = shopRows.map(r=>[r.period,r.name,r.id,Object.entries(r.byShop).map(([s,ms])=>`${s}: ${duration(ms)}`).join('\n'),fmt(r.shopGross),fmt(r.lunchShop),fmt(Math.max(0,r.shopGross-r.lunchShop)),fmt(r.canteen),fmt(r.holding+r.other),fmt(r.unknown)]);
  renderTable('shop-table',shopHeaders,shopRows.map(r=>`<tr><td class="datecell">${esc(r.period)}</td><td class="person">${esc(r.name)}</td><td class="small-muted">${esc(r.id)}</td><td class="shop-cell">${shopVisual(r)}</td><td class="num">${fmt(r.shopGross)}</td><td class="num small-muted">${fmt(r.lunchShop)}</td><td class="num accent">${fmt(Math.max(0,r.shopGross-r.lunchShop))}</td><td class="num">${fmt(r.canteen)}</td><td class="num">${fmt(r.holding+r.other)}</td><td class="num">${fmt(r.unknown)}</td></tr>`));
  const timeline = timelineRows(rows), types = {gate:'Завод',shop:'Цех',holding:'Территория',canteen:'Столовая',unknown:'Неясно',lunch:'Обед',other:'Другая зона'};
  const timelineHeaders = ['Дата смены','Сотрудник','Таб. №','Тип','Место / событие','Начало','Конец','Время','Строки Excel'];
  const timelineData = timeline.map(r=>[r.date,r.name,r.id,types[r.kind],r.zone,stamp(r.start),stamp(r.end),fmt(r.ms),r.sources]);
  const sizes = new Map();
  for (const r of timeline) { const key=`${r.employeeKey}|${r.date}`; sizes.set(key,(sizes.get(key)||0)+1); }
  renderTable('timeline-table',timelineHeaders,timeline.map((r,i)=>{
    const first=i===0||timeline[i-1].employeeKey!==r.employeeKey||timeline[i-1].date!==r.date;
    const shared=first?`<td class="datecell group-cell" rowspan="${sizes.get(`${r.employeeKey}|${r.date}`)}">${esc(r.date)}</td><td class="person group-cell" rowspan="${sizes.get(`${r.employeeKey}|${r.date}`)}">${esc(r.name)}</td><td class="small-muted group-cell" rowspan="${sizes.get(`${r.employeeKey}|${r.date}`)}">${esc(r.id)}</td>`:'';
    return `<tr class="${first?'group-start':''}">${shared}<td><span class="type-pill ${r.kind==='shop'?'shop':r.kind==='gate'?'gate':r.kind==='unknown'?'unknown':''}">${esc(types[r.kind])}</span></td><td>${esc(r.zone)}</td><td class="datecell">${stamp(r.start)}</td><td class="datecell">${stamp(r.end)}</td><td class="num">${fmt(r.ms)}</td><td class="source">${esc(r.sources)}</td></tr>`;
  }));
  $('summary').innerHTML = [[new Set(rows.map(r=>r.employeeKey)).size,'Сотрудников'],[rows.filter(r=>r.level==='red').length,'Смен с отклонениями'],[rows.filter(r=>r.level==='amber'||r.level==='gray').length,'Смен для уточнения'],[duration(rows.reduce((n,r)=>n+r.shopGross,0)),'Подтверждено в цехах']].map(([v,l])=>`<div class="metric"><b>${esc(v)}</b><span>${esc(l)}</span></div>`).join('');
  exportSheets = [
    {name:'План и факт',headers,rows:data,columnWidths:[18,30,14,28,26,26,22,16,18,16,60,48,55],rowStyles:rows.map(r=>styleFor(r.level)),cellStyles:rows.map(r=>[,,,,,,,,,,styleStrong(r.level)])},
    {name:'Цеха',headers:shopHeaders,rows:shopData,columnWidths:[18,30,14,48,16,18,16,16,18,18],rowStyles:shopRows.map((_,i)=>i%2?S.grayRow:S.plain),cellStyles:shopRows.map(r=>[,,,r.shopGross?S.green:undefined,undefined,r.lunchShop?S.amber:undefined,r.shopGross>r.lunchShop?S.green:undefined,r.canteen?S.amber:undefined,r.holding+r.other?S.blue:undefined,r.unknown?S.red:undefined])},
    {name:'Хронология',headers:timelineHeaders,rows:timelineData,columnWidths:[18,30,14,18,35,26,26,14,50],rowStyles:timeline.map(r=>r.kind==='unknown'?S.redRow:r.kind==='shop'?S.greenRow:r.kind==='gate'?S.blueRow:r.kind==='canteen'||r.kind==='lunch'?S.amberRow:S.grayRow)},
  ];
}

function refresh(resetDates = false) {
  $('workspace').hidden = !rawEvents.length && !schedule;
  $('clear-data').hidden = !rawEvents.length && !schedule && !presence;
  if ($('workspace').hidden) return;
  const sorted = dedupeEvents(rawEvents), lastDate = sorted.length ? localDate(sorted.at(-1).ts) : '';
  $('from-date').max = lastDate;
  $('to-date').max = lastDate;
  if (resetDates || !$('from-date').value) {
    if (schedule) {
      $('from-date').value = `${schedule.year}-${String(schedule.month).padStart(2,'0')}-01`;
      $('to-date').value = localDate(new Date(schedule.year,schedule.month,0,12).getTime());
      if (rawEvents.length) {
        const sorted=dedupeEvents(rawEvents), first=localDate(sorted[0].ts), last=localDate(sorted.at(-1).ts);
        if(first<=$('to-date').value && last>=$('from-date').value) {
          if(first>$('from-date').value)$('from-date').value=first;
          if(last<$('to-date').value)$('to-date').value=last;
        }
      }
    } else {
      const sorted=dedupeEvents(rawEvents);
      $('from-date').value=localDate(sorted[0].ts); $('to-date').value=localDate(sorted.at(-1).ts);
    }
  }
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
for (const id of ['calculate','sort','group']) $(id).addEventListener(id==='calculate'?'click':'change',()=>{try{recalculate();notify('');}catch(error){notify(error.message);}});
$('export-all').addEventListener('click',async()=>{try{recalculate();await downloadWorkbook(exportSheets,`PERCo_${$('from-date').value}_${$('to-date').value}.xlsx`);}catch(error){notify(`Не удалось создать Excel: ${error.message}`);}});
document.querySelectorAll('.tab').forEach(tab=>tab.addEventListener('click',()=>{
  document.querySelectorAll('.tab').forEach(x=>{const active=x===tab;x.classList.toggle('active',active);x.setAttribute('aria-selected',String(active));});
  document.querySelectorAll('[data-panel]').forEach(panel=>{panel.hidden=panel.dataset.panel!==tab.dataset.tab;});
}));
$('clear-data').addEventListener('click',()=>{
  if(!confirm('Удалить загруженные данные?'))return;
  rawEvents=[];eventFiles=[];schedule=null;presence=null;shiftReport=null;exportSheets=[];
  $('workspace').hidden=true;$('clear-data').hidden=true;
  for(const id of ['file-list','schedule-label','presence-label'])$(id).textContent='Файл не выбран';
  $('from-date').value='';$('to-date').value='';$('from-date').max='';$('to-date').max='';notify('Данные удалены.');
});
