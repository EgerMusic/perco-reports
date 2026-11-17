import { normalizeId, personKey, dedupeEvents, localDate, periodKey, OUTSIDE, HOLDING } from './core.mjs';

const HOUR = 3600000, MINUTE = 60000;
const clean = s => String(s ?? '').trim().toLocaleLowerCase('ru').replace(/ё/g, 'е').replace(/\s+/g, ' ');
export function shortName(name) {
  const parts = clean(name).split(/\s+/), surname = parts.shift() || '';
  return surname + ' ' + parts.map(p => p.replace(/[^а-яa-z]/g, '')).filter(Boolean).map(p => p.length <= 3 && parts.length === 1 ? p : p[0]).join('');
}
const ref = (fileName, rowNumber, sheetName = '') => `${fileName}${sheetName ? ` · лист ${sheetName}` : ''} · строка ${rowNumber}`;
const rowRef = (rows, fileName, ri) => ref(fileName, rows.rowSources?.[ri]?.rowNumber || ri + 1, rows.rowSources?.[ri]?.sheetName || rows.sheetName || '');
const onDate = (date, hour = 0, minute = 0) => { const d = new Date(`${date}T00:00:00`); d.setHours(hour, minute, 0, 0); return d.getTime(); };
const nextDate = date => { const d = new Date(`${date}T12:00:00`); d.setDate(d.getDate() + 1); return localDate(d.getTime()); };
const prevDate = date => { const d = new Date(`${date}T12:00:00`); d.setDate(d.getDate() - 1); return localDate(d.getTime()); };
const isShop = z => /^цех(?:\s|$)/i.test(z);
const normZone = z => z === 'Гостевой' ? HOLDING : z;
const kindOf = z => isShop(z) ? 'shop' : /столов/i.test(z) ? 'canteen' : normZone(z) === HOLDING ? 'holding' : z === OUTSIDE ? 'outside' : 'other';

export function parseSchedule(rows, fileName = '') {
  const months = ['январ','феврал','март','апрел','ма','июн','июл','август','сентябр','октябр','ноябр','декабр'];
  const titleRows = rows.filter((_, i) => rows.rowSources ? rows.rowSources[i].rowNumber <= 5 : i < 5);
  const title = titleRows.flat().join(' ').toLowerCase();
  const year = Number(title.match(/20\d{2}/)?.[0]);
  const month = months.findIndex(m => new RegExp(`(?:^|\\s)${m}`).test(title)) + 1;
  if (!year || !month) throw new Error('В заголовке графика должны быть месяц и год.');
  if (rows.rowSources) {
    const titles = new Map();
    rows.forEach((row, i) => { const src=rows.rowSources[i];if(src.rowNumber<=5)titles.set(src.sheetName,`${titles.get(src.sheetName)||''} ${row.join(' ')}`.toLowerCase()); });
    for (const text of titles.values()) {
      const sheetYear=Number(text.match(/20\d{2}/)?.[0]),sheetMonth=months.findIndex(m=>new RegExp(`(?:^|\\s)${m}`).test(text))+1;
      if(sheetYear&&sheetMonth&&(sheetYear!==year||sheetMonth!==month))throw new Error('В книге несколько месяцев. Загрузите график одного месяца отдельным файлом.');
    }
  }
  let columns = null;
  const people = [], warnings = [];
  rows.forEach((row, ri) => {
    const values = row.map(v => String(v ?? '').trim());
    if (values.includes('ФИО') && values.some(v => /^Таб\./i.test(v))) {
      columns = { name: values.indexOf('ФИО'), id: values.findIndex(v => /^Таб\./i.test(v)), days: values.map((v, i) => ({ day: Number(v), i })).filter(x => Number.isInteger(x.day) && x.day >= 1 && x.day <= 31) };
      if (!columns.days.length) throw new Error('В месячном графике не найдены столбцы с числами дней 1–31.');
      return;
    }
    if (!columns || !values[columns.id] || !values[columns.name]) return;
    const id = normalizeId(values[columns.id]), name = values[columns.name];
    if (!/^[\p{L}\d]+$/u.test(id)) return;
    const days = {};
    for (const { day, i } of columns.days) {
      const date = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
      if (new Date(`${date}T12:00:00`).getMonth() + 1 !== month) continue;
      const color = rows.cellStyles?.[ri]?.[i]?.fillColor || 'FFFFFF';
      const rgb = color.match(/../g)?.map(v => parseInt(v, 16)) || [];
      const gray = rgb.length === 3 && Math.max(...rgb) - Math.min(...rgb) <= 12 && rgb[0] >= 50 && rgb[0] < 225;
      const white = rgb.length === 3 && Math.min(...rgb) >= 235;
      if (rows.cellStyles?.[ri]?.[i]?.conditional) throw new Error('В графике есть условное форматирование дней: сохраните дневные/ночные цвета обычной заливкой ячеек.');
      days[date] = { code: clean(values[i]), night: gray, knownColor: gray || white, color, source: rowRef(rows, fileName, ri) };
    }
    people.push({ id, name, short: shortName(name), days, source: rowRef(rows, fileName, ri) });
  });
  if (!people.length) throw new Error('В месячном графике не найдены сотрудники.');
  const uniquePeople = new Set();
  for (const p of people) {
    const key = `${p.id}|${p.short}`;
    if (uniquePeople.has(key)) throw new Error(`Повтор сотрудника ${p.name}, таб. № ${p.id} в месячном графике: уберите дублирующую строку или уточните ФИО.`);
    uniquePeople.add(key);
  }
  const ids = new Map();
  for (const p of people) { if (!ids.has(p.id)) ids.set(p.id, []); ids.get(p.id).push(p); }
  for (const list of ids.values()) if (list.length > 1) for (const p of list) {
    p.conflict = 'Повтор табельного номера в месячном графике';
    warnings.push(`${p.name}: ${p.conflict} (${p.source})`);
  }
  return { people, year, month, fileName, warnings };
}

export function parsePresence(rows, fileName = '') {
  const hi = rows.findIndex(row => row?.map(v => String(v ?? '').trim()).includes('Табельный номер') && row?.map(v => String(v ?? '').trim()).includes('График работы'));
  if (hi < 0) throw new Error('Не найдены столбцы сводки присутствия.');
  let headers = rows[hi].map(v => String(v ?? '').trim());
  const c = n => headers.indexOf(n);
  if (['Дата','Табельный номер','Фамилия','График работы'].some(n=>c(n)<0)) throw new Error('В сводке отсутствуют обязательные столбцы даты, табельного номера, фамилии или графика.');
  const people = new Map();
  for (let i = hi + 1; i < rows.length; i++) {
    const row = rows[i] || [], trimmed=row.map(v=>String(v??'').trim());
    if(trimmed.includes('Табельный номер')&&trimmed.includes('График работы')){headers=trimmed;continue;}
    const id = normalizeId(row[c('Табельный номер')]);
    if (!id || !/^\d{2}[-.]\d{2}[-.]\d{4}$/.test(String(row[c('Дата')] || ''))) continue;
    const name = ['Фамилия','Имя','Отчество'].map(n => String(row[c(n)] ?? '').trim()).filter(Boolean).join(' ');
    const department = String(row[c('Подразделение')] ?? '').trim(), key = personKey(id, name, department);
    if (!people.has(key)) people.set(key, { id, name, short: shortName(name), department, templates: new Set(), source: rowRef(rows, fileName, i) });
    const value = String(row[c('График работы')] ?? '').trim();
    if (value) people.get(key).templates.add(value);
  }
  if (!people.size) throw new Error('В сводке не найдены сотрудники с датами и табельными номерами. Проверьте формат файла.');
  return { people: [...people.values()], fileName };
}

function profileFor(person, presence) {
  const matches = (presence?.people || []).filter(p => p.id === person.id && p.short === person.short && (!person.department || clean(p.department) === clean(person.department)));
  const templates = [...new Set(matches.flatMap(p => [...p.templates]))];
  if (templates.length > 1) return { conflict: 'Разные графики K для сотрудника', label: templates.join('; ') };
  if (!templates.length) return null;
  const label = templates[0], m = label.match(/(\d{1,2}):(\d{2}).*?(\d{1,2}):(\d{2})/);
  if (!m) return { conflict: `Не определены часы графика K: ${label}`, label };
  const start = +m[1] * 60 + +m[2], end = +m[3] * 60 + +m[4];
  if (start === 420 && end === 1140) return { start, end, rotating: true, label };
  if ((start === 420 && end === 960) || (start === 480 && end === 1020)) return { start, end, rotating: false, label };
  return { conflict: `Неподдерживаемый график K: ${label}`, label };
}

function expected(date, day, profile) {
  const hours = Number(day?.code), isWork = Number.isFinite(hours) && hours > 0;
  if (!isWork) {
    const knownOff = ['в','от','0'].includes(day?.code);
    return { isWork: false, knownOff, label: day?.code === 'в' ? 'Выходной' : day?.code === 'от' ? 'Отпуск' : day?.code || 'Нет плана', start: null, end: null, conflict: day?.code && !knownOff ? `Не распознано обозначение в плане: ${day.code}` : '' };
  }
  if (profile?.conflict) return { isWork: true, label: profile.label, start: null, end: null, conflict: profile.conflict };
  if (profile && !profile.rotating) {
    const start = onDate(date) + profile.start * MINUTE, end = onDate(date) + profile.end * MINUTE;
    return { isWork: true, start, end, label: profile.label, conflict: day.night ? 'Ночная ячейка при дневном графике K' : hours !== (profile.end - profile.start) / 60 - 1 ? 'Разные часы в месячном графике и K' : '' };
  }
  if (hours !== 11) return { isWork: true, start: null, end: null, label: `${day.code} ч по плану`, conflict: 'Не определены часы начала и конца этой смены' };
  if (!day.knownColor) return { isWork: true, start: null, end: null, label: '11 ч по плану', conflict: 'Не распознан цвет дневной/ночной смены' };
  return { isWork: true, start: onDate(date, day.night ? 19 : 7), end: onDate(day.night ? nextDate(date) : date, day.night ? 7 : 19), label: day.night ? 'Ночная 19:00–07:00' : 'Дневная 07:00–19:00' };
}

function slotFor(ts, profile) {
  let date = localDate(ts);
  const h = new Date(ts).getHours();
  // Fixed daytime templates never acquire a night schedule from observed passes.
  if (profile && !profile.rotating && !profile.conflict && h < profile.end / 60) return { date, start: onDate(date) + profile.start * MINUTE, end: onDate(date) + profile.end * MINUTE, night: false };
  const night = h >= 13 || h < 1;
  if (night && h < 1) date = prevDate(date);
  return { date, start: onDate(date, night ? 19 : 7), end: onDate(night ? nextDate(date) : date, night ? 7 : 19), night };
}

function episodes(events, profile) {
  const result = [];
  let current = null;
  for (const e of events) {
    const gateIn = e.from === OUTSIDE;
    if (gateIn) {
      if (current?.events.length) { current.broken = true; result.push(current); }
      current = { ...slotFor(e.ts, profile), events: [], gateIn: e };
    }
    if (!current) {
      // A leading morning exit normally closes the preceding night; it must not
      // contaminate the next evening's complete shift on the same calendar day.
      const leadingExit = e.to === OUTSIDE || isShop(e.from);
      const anchor = leadingExit && (!profile || profile.rotating) ? e.ts - 12 * HOUR : e.ts;
      current = { ...slotFor(anchor, profile), events: [], missingGate: true };
    }
    current.events.push(e);
    if (e.to === OUTSIDE) { current.gateOut = e; result.push(current); current = null; }
  }
  if (current?.events.length) result.push(current);
  return result;
}

function factOf(eps, date, profile) {
  const all = eps.flatMap(e => e.events).sort((a, b) => a.ts - b.ts);
  const entries = all.filter(e => isShop(e.to)), exits = all.filter(e => isShop(e.from));
  const first = entries[0] || null, last = exits.at(-1) || null;
  const byShop = {}, segments = [], messages = [], visits = [];
  let canteen = 0, holding = 0, unknown = 0, other = 0;
  for (const ep of eps) {
    if (ep.gateIn && ep.gateOut) visits.push({ start: ep.gateIn.ts, end: ep.gateOut.ts });
    if (ep.broken) messages.push('Повторный вход без выхода');
    if (ep.events.at(-1).ts - ep.events[0].ts > 20 * HOUR) messages.push('Интервал длиннее 20 часов: проверь отметки');
    for (let i = 0; i < ep.events.length - 1; i++) {
      const a = ep.events[i], b = ep.events[i + 1], ms = b.ts - a.ts;
      if (ms <= 0) continue;
      const zone = normZone(a.to), coherent = zone === normZone(b.from);
      const kind = coherent ? kindOf(zone) : 'unknown';
      // A gap of more than one normal shift is not confirmed shop time.
      if (!coherent || ms > 20 * HOUR) { unknown += ms; segments.push({ start: a.ts, end: b.ts, zone: 'Не распределено', kind: 'unknown', a, b }); continue; }
      if (kind === 'shop') byShop[zone] = (byShop[zone] || 0) + ms;
      else if (kind === 'canteen') canteen += ms;
      else if (kind === 'holding') holding += ms;
      else if (kind !== 'outside') other += ms;
      if (kind !== 'outside') segments.push({ start: a.ts, end: b.ts, zone, kind, a, b });
    }
  }
  const span = first && last && last.ts > first.ts ? last.ts - first.ts : null;
  const startGood = Boolean(first && !exits.some(e => e.ts < first.ts));
  const boundaryGood = Boolean(startGood && last && last.ts > first.ts && !entries.some(e => e.ts > last.ts));
  const shopGross = Object.values(byShop).reduce((a, b) => a + b, 0);
  const lunchShop = Math.min(shopGross, Math.max(0, HOUR - Math.min(canteen, HOUR)));
  const sorted = [...visits].sort((a, b) => a.start - b.start);
  const gaps = sorted.slice(1).map((v, i) => Math.max(0, v.start - sorted[i].end)).filter(Boolean);
  return { first, last, span, startGood, boundaryGood, byShop, shopGross, lunchShop, canteen, holding, unknown, other, segments, visits, gaps, messages, observed: all.length > 0, all, date, night: eps[0]?.night };
}

export function buildShiftReport(events, schedule, presence, fromDate, toDate) {
  if (!fromDate || !toDate || fromDate > toDate) throw new Error('Проверь даты отчётного периода.');
  const sorted = dedupeEvents(events), actualPeople = new Map(), people = [], rows = [];
  for (const e of sorted) {
    const key = e.employeeKey || personKey(e.id, e.name, e.department || '');
    if (!actualPeople.has(key)) actualPeople.set(key, { id: e.id, name: e.name, short: shortName(e.name), department: e.department || '', key, events: [] });
    actualPeople.get(key).events.push(e);
  }
  const used = new Set();
  for (const plan of schedule?.people || []) {
    const candidates = [...actualPeople.values()].filter(p => p.id === plan.id && p.short === plan.short);
    const actual = candidates.length === 1 ? candidates[0] : null;
    if (actual) used.add(actual.key);
    people.push({ ...plan, ...actual, plan, key: actual?.key || `plan|${plan.id}|${plan.short}`, events: actual?.events || [], matchConflict: candidates.length > 1 ? 'Неоднозначное ФИО/подразделение' : plan.conflict || '' });
  }
  for (const actual of actualPeople.values()) if (!used.has(actual.key)) people.push(actual);
  const coverageStart = sorted[0]?.ts ?? Infinity, coverageEnd = sorted.at(-1)?.ts ?? -Infinity;
  // A monthly plan is not evidence for dates beyond the uploaded passes.
  if (!sorted.length) return { rows, shops: [], coverageStart, coverageEnd, warnings: schedule?.warnings || [] };
  const throughDate = localDate(coverageEnd);
  if (toDate > throughDate) toDate = throughDate;
  const coverageRanges = [...new Map(events.filter(e=>e.fileCoverage).map(e=>[`${e.fileCoverage.start}|${e.fileCoverage.end}`,e.fileCoverage])).values()].sort((a,b)=>a.start-b.start);
  const covers = (start,end) => {
    if(start==null)return false;
    if(!coverageRanges.length)return coverageStart<=start&&coverageEnd>=end;
    let cursor=start;
    for(const range of coverageRanges){if(range.end<cursor)continue;if(range.start>cursor)return false;cursor=Math.max(cursor,range.end);if(cursor>=end)return true;}
    return false;
  };
  for (const person of people) {
    const profile = profileFor(person, presence), eps = episodes(person.events, profile);
    const slots = new Map();
    for (const ep of eps) {
      // Reentries after midnight and daytime reentries belong to the original shift.
      const previous = [...slots.values()].flat().at(-1);
      const nextPlan = expected(ep.date, person.plan?.days[ep.date], profile);
      const newPlannedShift = previous && ep.start !== previous.start && nextPlan.start === ep.start && ep.events[0].ts >= ep.start - 3 * HOUR;
      const previousFinished = previous?.gateOut && previous.gateOut.ts >= previous.end - HOUR;
      if (previous && ep.start !== previous.start && !newPlannedShift && !previousFinished && ep.events[0].ts <= previous.end + HOUR && ep.events[0].ts >= previous.start - 3 * HOUR) {
        ep.date = previous.date; ep.start = previous.start; ep.end = previous.end; ep.night = previous.night;
      }
      if (!slots.has(ep.start)) slots.set(ep.start, []);
      slots.get(ep.start).push(ep);
    }
    const dates = new Set([...Object.keys(person.plan?.days || {}), ...eps.map(ep => ep.date)]);
    for (const date of [...dates].sort()) {
      if (date < fromDate || date > toDate) continue;
      const day = person.plan?.days[date], planned = expected(date, day, profile);
      const groups = [...slots.values()].filter(es => es[0].date === date);
      if (!planned.isWork && !groups.length) continue;
      const jobs = [];
      if (!groups.length) jobs.push({ plan: planned, actual: [] });
      else if (planned.isWork && groups.length > 1) {
        const plannedNight = planned.start != null ? new Date(planned.start).getHours() >= 13 : null;
        const matched = groups.find(es => es[0].night === plannedNight);
        if (matched) jobs.push({ plan: planned, actual: matched });
        else jobs.push({ plan: planned, actual: [] });
        for (const actual of groups) if (actual !== matched) jobs.push({ plan: { isWork: false, knownOff: true, label: 'Вне плановой смены', start: null, end: null }, actual });
      } else for (const actual of groups) jobs.push({ plan: planned, actual });
      for (const { plan, actual } of jobs) {
      const fact = factOf(actual, date, profile), messages = [], state = { level: 'green', label: 'Вышел в свою смену' };
      const warn = (text, level = 'amber') => { messages.push(text); if (level === 'red' || state.level === 'green') state.level = level; };
      if (!person.plan) { state.label = 'Нет месячного плана'; state.level = 'gray'; }
      if (person.matchConflict) warn(person.matchConflict);
      if (plan.conflict) warn(plan.conflict);
      const covered = covers(plan.start,plan.end);
      if (!fact.observed) {
        const missingShift = covered && person.events.length > 0 && !person.matchConflict && !plan.conflict;
        state.level = missingShift ? 'red' : 'gray';
        state.label = missingShift ? 'Нет отметок за плановую смену' : 'Недостаточно данных';
        if (!person.events.length) messages.push('Сотрудник не найден в подробных проходах');
        else if (!covered) messages.push('Смена не покрыта выгрузкой полностью');
      } else {
        if (plan.knownOff) { state.label = 'Выход вне графика'; warn(plan.label, 'red'); }
        else if (day && !plan.isWork) { state.label = 'Не определён план смены'; warn(plan.conflict || 'Нет обозначения смены на эту дату'); }
        if (plan.isWork && plan.start != null && fact.first && fact.startGood) {
          const plannedNight = new Date(plan.start).getHours() >= 13;
          const observedNight = new Date(fact.first.ts).getHours() >= 13;
          if (plannedNight !== observedNight) { state.label = 'Другая смена'; warn('Фактическая смена отличается от плановой', 'red'); }
          else {
            // Compare to the minute, matching the precision of the approved schedule.
            const late = Math.floor(fact.first.ts / MINUTE) - Math.floor(plan.start / MINUTE);
            const early = fact.last ? Math.floor(plan.end / MINUTE) - Math.floor(fact.last.ts / MINUTE) : 0;
            if (late > 0) warn(`Поздний вход в цех: ${late} мин`, 'red');
            if (early > 0 && covered && fact.boundaryGood) warn(`Ранний выход из цеха: ${early} мин`, 'red');
          }
        }
        if (!fact.boundaryGood || fact.messages.length || !covered && person.plan && plan.isWork) {
          warn('Недостаточно данных для полного расчёта смены');
          state.label = 'Недостаточно данных';
        }
        if (!fact.first && !fact.last) { state.label = 'Нет подтверждённых проходов в цех'; state.level = 'amber'; }
        if (fact.unknown) warn('Есть нераспределённые интервалы');
        for (const text of fact.messages) warn(text);
      }
      rows.push({ id: person.id, name: person.name, employeeKey: person.key, department: person.department || '', period: date, date, ts: actual[0]?.start || plan.start || onDate(date), plan, day, profile, ...fact, status: state.label, level: state.level, comments: [...new Set(messages)], planSource: day?.source || '', sources: fact.all.map(e => ref(e.fileName, e.rowNumber, e.sheetName)) });
      }
    }
  }
  return { rows, shops: [...new Set(rows.flatMap(r => Object.keys(r.byShop)))].sort(), coverageStart, coverageEnd, warnings: schedule?.warnings || [] };
}

export function aggregateShops(report, granularity) {
  const groups = new Map();
  for (const row of report.rows) {
    const period = periodKey(row.ts, granularity), key = `${row.employeeKey}|${period}`;
    if (!groups.has(key)) groups.set(key, { id: row.id, name: row.name, employeeKey: row.employeeKey, period, byShop: {}, shopGross: 0, lunchShop: 0, canteen: 0, holding: 0, other: 0, unknown: 0 });
    const g = groups.get(key);
    for (const field of ['shopGross','lunchShop','canteen','holding','other','unknown']) g[field] += row[field];
    for (const [shop, ms] of Object.entries(row.byShop)) g.byShop[shop] = (g.byShop[shop] || 0) + ms;
  }
  return { rows: [...groups.values()], shops: report.shops };
}
