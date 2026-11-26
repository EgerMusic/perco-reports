import { normalizeId, personKey, dedupeEvents, parseStamp, localDate, periodKey, OUTSIDE, HOLDING } from './core.mjs';

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
const hasShopMarks = group => group.some(ep => ep.events.some(e => isShop(e.from) || isShop(e.to)));
const normZone = z => z === 'Гостевой' ? HOLDING : z;
const kindOf = z => isShop(z) ? 'shop' : /столов/i.test(z) ? 'canteen' : normZone(z) === HOLDING ? 'holding' : z === OUTSIDE ? 'outside' : 'other';
const displayDate = date => date.split('-').reverse().join('.');
const clock = ts => new Date(ts).toTimeString().slice(0, 5);
const moment = (ts, date) => `${localDate(ts) === date ? '' : `${displayDate(localDate(ts))} `}${clock(ts)}`;
const shiftHours = ms => `${Math.floor(ms / HOUR)} ч${ms % HOUR ? ` ${Math.floor(ms % HOUR / MINUTE)} мин` : ''}`;
const windowText = ({start, end}, date) => `${moment(start, date)}–${moment(end, date)}`;
const identity = p => `${p.id}|${p.short}`;

function indexPeople(people, keyOf = identity) {
  const index = new Map();
  for (const person of people) {
    const key = keyOf(person);
    if (!index.has(key)) index.set(key, []);
    index.get(key).push(person);
  }
  return index;
}

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

export function mergeSchedules(schedules) {
  if (!schedules.length) return null;
  const peopleByIdentity = new Map();
  const warnings = new Set(schedules.flatMap(s => s.warnings));
  for (const schedule of schedules) for (const person of schedule.people) {
    const key = identity(person);
    if (!peopleByIdentity.has(key)) {
      peopleByIdentity.set(key, { ...person, days: {}, dayChoices: new Map(), fullNames: new Set() });
    }
    const merged = peopleByIdentity.get(key);
    const fullName = clean(person.name).split(' ').slice(1).filter(part => /^\p{L}{2,}$/u.test(part)).length >= 2;
    if (fullName) merged.fullNames.add(clean(person.name));
    for (const [date, day] of Object.entries(person.days)) {
      if (!merged.dayChoices.has(date)) merged.dayChoices.set(date, new Map());
      const choices = merged.dayChoices.get(date);
      const shift = Number(day.code) === 11 ? day.knownColor ? day.night ? 'ночная' : 'дневная' : 'цвет не распознан' : '';
      const value = shift ? `${day.code} (${shift})` : day.code;
      if (!choices.has(value)) choices.set(value, { day, sources: new Set() });
      choices.get(value).sources.add(day.source);
    }
  }
  const people = [...peopleByIdentity.values()].map(({dayChoices, fullNames, ...person}) => {
    if (fullNames.size > 1) {
      person.conflict = `Для таб. № ${person.id} в месячных графиках указаны разные полные ФИО с одинаковыми инициалами.`;
      warnings.add(person.conflict);
    }
    for (const [date, choices] of dayChoices) {
      if (choices.size > 1) choices.delete('');
      const [{ day }] = choices.values();
      const source = [...new Set([...choices.values()].flatMap(choice => [...choice.sources]))].join('\n');
      if (choices.size === 1) {
        person.days[date] = { ...day, source };
      } else {
        const variants = [...choices].map(([value, choice]) => `«${value}» (${[...choice.sources].join(', ')})`);
        const conflict = `На ${displayDate(date)} месячные графики расходятся. ${variants.join(' / ')}.`;
        person.days[date] = { code:'', conflict, source };
        warnings.add(`${person.name}. ${conflict}`);
      }
    }
    return person;
  });
  for (const group of indexPeople(people, p => p.id).values()) if (group.length > 1) {
    for (const person of group) {
      person.conflict ||= 'Повтор табельного номера в месячных графиках';
      warnings.add(`${person.name}. ${person.conflict} (${person.source}).`);
    }
  }
  const ordered = [...schedules].sort((a,b) => a.year - b.year || a.month - b.month);
  const first = ordered[0], last = ordered.at(-1);
  return {
    people, warnings: [...warnings],
    fromDate: localDate(new Date(first.year, first.month - 1, 1, 12).getTime()),
    toDate: localDate(new Date(last.year, last.month, 0, 12).getTime()),
  };
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
    const dateText = String(row[c('Дата')] || '').trim(), dateStamp = parseStamp(`${dateText} 00:00:00`);
    if (!id || !Number.isFinite(dateStamp)) continue;
    const date = localDate(dateStamp);
    const name = ['Фамилия','Имя','Отчество'].map(n => String(row[c(n)] ?? '').trim()).filter(Boolean).join(' ');
    const department = String(row[c('Подразделение')] ?? '').trim(), key = personKey(id, name, department);
    if (!people.has(key)) people.set(key, { id, name, short: shortName(name), department, templates: new Set(), source: rowRef(rows, fileName, i), events: [], days: {} });
    const person = people.get(key);
    const value = String(row[c('График работы')] ?? '').trim();
    if (value) person.templates.add(value);
    const record = { date, template: value, source: rowRef(rows, fileName, i), issues: [], presence: String(row[c('Присутствие')] ?? '').trim() };
    record.hasPresence = /^(\d+):([0-5]\d)(?::([0-5]\d))?$/.test(record.presence) && record.presence.split(':').some(part => Number(part) > 0);
    if (!person.days[date]) person.days[date] = [];
    person.days[date].push(record);
    for (const [column, from, to] of [['Вход', OUTSIDE, HOLDING], ['Выход', HOLDING, OUTSIDE]]) {
      const clock = String(row[c(column)] ?? '').trim();
      if (!clock || /^[-—]+$/.test(clock)) continue;
      let ts;
      if (/^\d{1,2}:\d{2}(?::\d{2})?$/.test(clock)) ts = parseStamp(`${dateText} ${clock.padStart(clock.length <= 5 ? 5 : 8, '0')}`);
      else ts = parseStamp(clock);
      if (!Number.isFinite(ts) || localDate(ts) !== date) { record.issues.push(`Не распознано время «${column}»: ${clock}`); continue; }
      const event = { id, name, department, employeeKey: key, ts, from, to, fileName, rowNumber: rows.rowSources?.[i]?.rowNumber || i + 1, sheetName: rows.rowSources?.[i]?.sheetName || rows.sheetName || '', sourceType: 'presence', key: [key, ts, from, to, 'presence'].join('|') };
      person.events.push(event);
    }
  }
  if (!people.size) throw new Error('В сводке не найдены сотрудники с датами и табельными номерами. Проверьте формат файла.');
  for (const person of people.values()) person.events = dedupeEvents(person.events);
  const times = [...people.values()].flatMap(p => p.events.map(e => e.ts));
  const coverage = times.length ? times.reduce((r, ts) => ({ start: Math.min(r.start, ts), end: Math.max(r.end, ts) }), { start: Infinity, end: -Infinity }) : null;
  return { people: [...people.values()], fileName, coverage };
}

export function reportCoverage(events, presence) {
  const bounds = events.reduce(({ start, end }, { ts }) => ({ start:Math.min(start, ts), end:Math.max(end, ts) }), { start:Infinity, end:-Infinity });
  if (presence?.coverage) {
    bounds.start = Math.min(bounds.start, presence.coverage.start);
    bounds.end = Math.max(bounds.end, presence.coverage.end);
  }
  return Number.isFinite(bounds.end) ? bounds : null;
}

function profileFor(matches, date) {
  const dateRows = matches.flatMap(p => p.days[date] || []).filter(record => record.template);
  const templates = [...new Set(dateRows.length ? dateRows.map(r => r.template) : matches.flatMap(p => [...p.templates]))];
  if (!templates.length) return null;
  const source = [...new Set(dateRows.length ? dateRows.map(r => r.source) : matches.map(p => p.source))].join('\n');
  if (templates.length > 1) return { conflict: `В столбце K несколько графиков (${templates.join(' / ')}). Уточни, какие часы действуют ${displayDate(date)}.`, label: templates.join(' / '), source };
  const label = templates[0], m = label.match(/(\d{1,2}):(\d{2}).*?(\d{1,2}):(\d{2})/);
  if (!m || +m[1] > 23 || +m[3] > 23 || +m[2] > 59 || +m[4] > 59) return { conflict: `В столбце K указано «${label}». По этому значению нельзя определить часы входа и выхода.`, label, source };
  const start = +m[1] * 60 + +m[2], end = +m[3] * 60 + +m[4];
  if (start === end) return { conflict: `В столбце K указано «${label}». Время начала и конца совпадает.`, label, source };
  return { start, end, rotating: start === 420 && end === 1140, night: end < start, label, source };
}

function normFor(date, profile, night) {
  if (profile.conflict) return { start:null, end:null, label:profile.label, conflict:profile.conflict, normSource:profile.source };
  if (profile.rotating && night) {
    return { start:onDate(date, 19), end:onDate(nextDate(date), 7), label:'Ночная 19:00–07:00', night:true, normSource:profile.source };
  }
  return { start:onDate(date) + profile.start * MINUTE, end:onDate(profile.night ? nextDate(date) : date) + profile.end * MINUTE, label:profile.label, night:profile.night, normSource:profile.source };
}

function expected(date, day, profile, observedNight) {
  if (day?.conflict) {
    return { isWork:false, unknownCode:true, label:'Расхождение графиков', start:null, end:null, conflict:day.conflict };
  }
  if (!day?.code && profile) return { ...normFor(date, profile, observedNight), isWork:false, checksTime:true, normOnly:true };
  const hours = Number(day?.code), isWork = Number.isFinite(hours) && hours > 0;
  if (!isWork) {
    const knownOff = ['в','от','0'].includes(day?.code);
    const unknownCode = Boolean(day?.code && !knownOff);
    return { isWork: false, knownOff, unknownCode, label: day?.code === 'в' ? 'Выходной' : day?.code === 'от' ? 'Отпуск' : day?.code || 'Нет плана', start: null, end: null, conflict: unknownCode ? `В месячном графике на ${displayDate(date)} указано «${day.code}». Это обозначение не распознано, поэтому неясно, была ли назначена смена.` : '' };
  }
  const monthlyWindow = hours === 11 && day.knownColor
    ? { start:onDate(date, day.night ? 19 : 7), end:onDate(day.night ? nextDate(date) : date, day.night ? 7 : 19) }
    : null;
  if (profile) {
    const norm = normFor(date, profile, day.night);
    const plan = { ...norm, isWork:true, monthlyWindow };
    if (profile.conflict) return plan;
    if (hours !== (norm.end - norm.start) / HOUR - 1 || monthlyWindow && (norm.start !== monthlyWindow.start || norm.end !== monthlyWindow.end)) {
      const monthly = monthlyWindow ? `${windowText(monthlyWindow, date)}, ${shiftHours(monthlyWindow.end - monthlyWindow.start)} с обедом` : `${day.code} ч по табелю`;
      plan.conflict = `Месячный график: ${monthly}. Столбец K: ${windowText(norm, date)}, ${shiftHours(norm.end - norm.start)}. Вход и выход проверяются по K.`;
      plan.hoursConflict = true;
    } else if (hours === 11 && profile.rotating && !day.knownColor) {
      return { ...plan, start:null, end:null, conflict:'По цвету ячейки месячного графика не удалось определить, дневная или ночная смена.' };
    }
    return plan;
  }
  if (monthlyWindow) return { ...monthlyWindow, isWork:true, night:day.night, label:day.night ? 'Ночная 19:00–07:00' : 'Дневная 07:00–19:00', normSource:day.source };
  const conflict = hours === 11 ? 'По цвету ячейки месячного графика не удалось определить, дневная или ночная смена.' : `В месячном графике указано ${day.code} ч, а в K нет часов для сравнения входа и выхода.`;
  return { isWork:true, start:null, end:null, label:`${day.code} ч по плану`, conflict };
}

function planForActual(date, day, profile, plan, group) {
  const night = group[0].night;
  if (plan.normOnly) return expected(date, day, profile, night);
  const rotating = profile ? profile.rotating : Number(day?.code) === 11 && day.knownColor;
  if (rotating && plan.isWork && plan.start != null && !plan.night && night &&
      group.some(ep => ep.gateIn || ep.events.some(e => isShop(e.to)))) {
    return expected(date, { ...day, night:true }, profile, true);
  }
  return plan;
}

function slotFor(ts, profile) {
  let date = localDate(ts);
  const h = new Date(ts).getHours();
  // Fixed daytime templates never acquire a night schedule from observed passes.
  if (profile && !profile.rotating && !profile.conflict && !profile.night && h < profile.end / 60) return { date, ...normFor(date, profile, false) };
  const night = h >= 13 || h < 1;
  if (night && h < 1) date = prevDate(date);
  if (profile?.night && night) return { date, ...normFor(date, profile, true) };
  return { date, start: onDate(date, night ? 19 : 7), end: onDate(night ? nextDate(date) : date, night ? 7 : 19), night };
}

function episodes(events, profileAt) {
  const result = [];
  let current = null;
  for (const e of events) {
    const profile = profileAt(localDate(e.ts));
    const gateIn = e.from === OUTSIDE;
    if (gateIn) {
      if (current?.events.length) { current.broken = true; result.push(current); }
      current = { ...slotFor(e.ts, profile), events: [], gateIn: e };
    }
    if (!current) {
      // A leading morning exit normally closes the preceding night; it must not
      // contaminate the next evening's complete shift on the same calendar day.
      const leadingExit = e.to === OUTSIDE || isShop(e.from);
      const anchor = leadingExit && (!profile || profile.rotating || profile.night) ? e.ts - 12 * HOUR : e.ts;
      current = { ...slotFor(anchor, profile), events: [], missingGate: true };
    }
    current.events.push(e);
    if (e.to === OUTSIDE) { current.gateOut = e; result.push(current); current = null; }
  }
  if (current?.events.length) result.push(current);
  return result;
}

function shiftGroups(events, profileAt, planDays, fromDate, keepIsolatedExit) {
  const eps = episodes(events, profileAt);
  const slots = new Map();
  const datesWithEpisode = new Set(eps.map(ep => ep.date));
  let lastGroup;
  for (const ep of eps) {
    // Keep an isolated exit at the left edge even when the guessed night falls outside the report.
    const markDate = localDate(ep.events[0].ts);
    if (ep.missingGate && ep.date < fromDate && markDate === fromDate && !datesWithEpisode.has(markDate) &&
        (keepIsolatedExit || ep.gateOut && !hasShopMarks([ep]))) {
      ep.date = markDate;
      ep.start = onDate(markDate);
    }
    const previous = lastGroup?.at(-1);
    const nextPlan = expected(ep.date, planDays[ep.date], profileAt(ep.date), ep.night);
    const newPlannedShift = previous && ep.start !== previous.start && nextPlan.start === ep.start && ep.events[0].ts >= ep.start - 3 * HOUR;
    const previousFinished = previous?.gateOut && previous.gateOut.ts >= previous.end - HOUR;
    if (previous && ep.start !== previous.start && !newPlannedShift && !previousFinished && ep.events[0].ts <= previous.end + HOUR && ep.events[0].ts >= previous.start - 3 * HOUR) {
      ep.date = previous.date; ep.start = previous.start; ep.end = previous.end; ep.night = previous.night;
    }
    if (!slots.has(ep.start)) {
      lastGroup = [];
      slots.set(ep.start, lastGroup);
    }
    slots.get(ep.start).push(ep);
  }
  const groupsByDate = new Map();
  for (const group of slots.values()) {
    const { date } = group[0];
    if (!groupsByDate.has(date)) groupsByDate.set(date, []);
    groupsByDate.get(date).push(group);
  }
  return groupsByDate;
}

function factOf(eps, date, useGate = false) {
  const all = eps.flatMap(e => e.events).sort((a, b) => a.ts - b.ts);
  const entries = all.filter(e => useGate ? e.from === OUTSIDE : isShop(e.to));
  const exits = all.filter(e => useGate ? e.to === OUTSIDE : isShop(e.from));
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

function factoryFactOf(eps, date) {
  const all = eps.flatMap(e => e.events).sort((a, b) => a.ts - b.ts);
  const entries = all.filter(e => e.from === OUTSIDE), exits = all.filter(e => e.to === OUTSIDE);
  const first = entries[0] || null, last = exits.at(-1) || null;
  const visits = eps.filter(ep => ep.gateIn && ep.gateOut && ep.gateOut.ts > ep.gateIn.ts).map(ep => ({ start: ep.gateIn.ts, end: ep.gateOut.ts, a: ep.gateIn, b: ep.gateOut }));
  const startGood = Boolean(first && !exits.some(e => e.ts < first.ts));
  const boundaryGood = Boolean(startGood && last && last.ts > first.ts && !entries.some(e => e.ts > last.ts));
  const messages = [];
  if (eps.some(ep => ep.broken)) messages.push('Повторный вход без выхода');
  if (eps.some(ep => ep.events.at(-1).ts - ep.events[0].ts > 20 * HOUR)) messages.push('Интервал длиннее 20 часов: проверь отметки');
  const gaps = visits.slice(1).map((v, i) => Math.max(0, v.start - visits[i].end)).filter(Boolean);
  return { first, last, span: first && last && last.ts > first.ts ? last.ts - first.ts : null, startGood, boundaryGood, byShop: {}, shopGross: 0, lunchShop: 0, canteen: 0, holding: 0, unknown: 0, other: 0, segments: [], visits, gaps, messages, observed: all.length > 0, all, date, night: eps[0]?.night };
}

function mergeSummaryPeople(people) {
  return [...indexPeople(people).values()].flatMap(matches => {
    if (matches.length === 1) return matches;
    // Transfers between departments can produce separate records for the same employee.
    const dates = matches.flatMap(p => Object.keys(p.days));
    const sameName = matches.every(p => clean(p.name) === clean(matches[0].name));
    if (!sameName || new Set(dates).size !== dates.length) return matches;
    return [{ ...matches[0], department:[...new Set(matches.map(p => p.department).filter(Boolean))].join(' / '),
      days:Object.assign({}, ...matches.map(p => p.days)), templates:new Set(matches.flatMap(p => [...p.templates])),
      events:dedupeEvents(matches.flatMap(p => p.events)) }];
  });
}

function reportPeople(events, schedule, presence) {
  const detailPeople = new Map();
  for (const e of events) {
    const key = e.employeeKey || personKey(e.id, e.name, e.department || '');
    if (!detailPeople.has(key)) detailPeople.set(key, { id:e.id, name:e.name, short:shortName(e.name), department:e.department || '', key, events:[] });
    detailPeople.get(key).events.push(e);
  }
  const originalSummaries = indexPeople(presence?.people || []);
  const summaries = indexPeople(mergeSummaryPeople(presence?.people || []));
  const details = indexPeople(detailPeople.values());
  const people = [...detailPeople.values()].map(person => {
    const variants = originalSummaries.get(identity(person)) || [];
    const sameDepartment = variants.filter(p => !person.department || clean(p.department) === clean(person.department));
    // A unique match on both sides can also use K when the department label changed.
    const matches = variants.length === 1 && details.get(identity(person)).length === 1 ? variants : sameDepartment;
    return { ...person, matches, ambiguousSummary:variants.length > 0 && (matches.length !== 1 || details.get(identity(person)).length !== 1), factSource:'Подробные проходы' };
  });
  for (const [key, matches] of summaries) {
    if (details.has(key)) continue;
    for (const person of matches) {
      people.push({ ...person, key:personKey(person.id, person.name, person.department), summary:person, matches:[person], factSource:'Сводка PERCo' });
    }
  }

  const candidates = indexPeople(people);
  const monthlyPeople = schedule?.people || [];
  const monthlyByName = indexPeople(monthlyPeople, p => p.short);
  const passesByName = indexPeople(people, p => p.short);
  for (const plan of monthlyPeople) {
    const matches = candidates.get(identity(plan)) || [];
    if (matches.length === 1) {
      Object.assign(matches[0], { plan, matchConflict:plan.conflict });
      continue;
    }
    let matchConflict = plan.conflict;
    if (matches.length > 1) {
      matchConflict = `По табельному номеру и инициалам найдено несколько записей (${matches.map(p => p.department || 'без подразделения').join(' / ')}). Нельзя однозначно связать месячный график с проходами.`;
      for (const person of matches) person.matchConflict = matchConflict;
    }
    people.push({ ...plan, plan, key:`plan|${identity(plan)}`, events:[], matches:[], matchConflict, factSource:'Только месячный график' });
  }

  for (const person of people) {
    const planOnly = person.factSource === 'Только месячный график';
    if ((person.plan && !planOnly) || !person.short.trimEnd().includes(' ')) continue;
    const otherSource = planOnly ? passesByName : monthlyByName;
    const namesakes = otherSource.get(person.short);
    if (!namesakes || namesakes.some(p => p.id === person.id)) continue;
    const otherNumbers = [...new Set(namesakes.map(p => `№ ${p.id}`))].join(', ');
    const monthlyNumbers = planOnly ? `№ ${person.id}` : otherNumbers;
    const passNumbers = planOnly ? otherNumbers : `№ ${person.id}`;
    person.numberWarning = `Проверьте табельные номера. При совпадающих фамилии и инициалах в месячном графике — ${monthlyNumbers}, в выгрузках проходов — ${passNumbers}.`;
  }
  return people;
}

function daysOfWindow(start, end) {
  if (start == null || end == null || start >= end) return [];
  const dates = [];
  const last = localDate(end - 1);
  for (let day = localDate(start); day <= last; day = nextDate(day)) dates.push(day);
  return dates;
}

function coverageNote(start, end, range, date, gapDays) {
  if (start == null) return 'Часы смены не определены. Проверить полное покрытие смены выгрузкой нельзя.';
  if (!range || !Number.isFinite(range.end)) return 'В этом источнике нет отметок за выбранный период.';
  if (range.end < end) return `Последняя отметка в выгрузке — ${displayDate(localDate(range.end))} в ${clock(range.end)}. Для проверки нужна выгрузка до конца смены (${moment(end, date)}).`;
  if (range.start > start) return `Первая отметка в выгрузке — ${displayDate(localDate(range.start))} в ${clock(range.start)}. Для проверки нужна выгрузка с начала смены (${moment(start, date)}).`;
  if (gapDays.length) return `В выгрузке нет ни одной отметки за ${gapDays.map(displayDate).join(', ')}. Возможно, данные за этот день пропущены.`;
  return 'В выгрузке есть разрыв или нет строки сотрудника за одну из дат смены.';
}

function attendanceResult({ person, plan, day, date, fact, accounting, sourceRange, sourceDays, covers, gateFallback, summaryRecord }) {
  const factory = accounting === 'factory';
  const entry = factory ? 'вход на завод' : 'вход в цех';
  const exit = factory ? 'выход с завода' : 'выход из цеха';
  const comments = [];
  let needsReview = false;
  let level = 'green', status = plan.normOnly ? 'По норме K' : 'Вышел в свою смену';
  const warn = message => {
    if (message) comments.push(message);
    needsReview = true;
    if (level === 'green') level = 'amber';
  };
  const deviation = (label, message) => {
    level = 'red'; status = label; comments.push(message);
  };
  if (person.matchConflict) warn(person.matchConflict);
  if (person.numberWarning) warn(person.numberWarning);
  if (plan.conflict) warn(plan.conflict);
  if (person.ambiguousSummary) warn('Не удалось однозначно связать сотрудника с записями сводки PERCo.');
  const finish = () => {
    if (needsReview) comments.push('Нужно ручное рассмотрение.');
    return { status, level, needsReview, comments:[...new Set(comments)] };
  };

  // For a no-show, check both possible windows when K and the roster disagree.
  const checkStart = !fact.observed && plan.monthlyWindow && plan.start != null ? Math.min(plan.start, plan.monthlyWindow.start) : plan.start;
  const checkEnd = !fact.observed && plan.monthlyWindow && plan.end != null ? Math.max(plan.end, plan.monthlyWindow.end) : plan.end;
  const windowDates = daysOfWindow(checkStart, checkEnd);
  const gapDays = windowDates.filter(d => !sourceDays.has(d));
  const sourceDates = [...new Set([date, ...windowDates, ...fact.all.map(e => localDate(e.ts))])];
  const summaryRows = sourceDates.flatMap(d => summaryRecord?.days[d] || []);
  const issues = summaryRows.flatMap(r => r.issues.map(issue => `${issue} (${r.source}).`));
  for (const issue of issues) warn(issue);
  if (gateFallback && fact.observed) {
    status = fact.first ? 'Был на заводе' : 'Зафиксирован выход с завода';
    level = 'gray';
    const origin = gateFallback === 'summary' ? 'из сводки PERCo' : 'из подробной выгрузки';
    warn(`За эту смену цеховых отметок нет. Показано время проходной ${origin}.`);
    if (plan.knownOff) {
      const offSchedule = `В месячном графике на ${displayDate(date)} — ${plan.label.toLowerCase()}.`;
      if (gateFallback === 'detail') deviation('Выход вне графика', offSchedule);
      else warn(offSchedule);
    }
    if (!fact.first) warn('Нет отметки входа на завод.');
    if (!fact.last) warn('Нет отметки выхода с завода.');
    if (fact.first && fact.last && !fact.boundaryGood) warn('Отметки проходной не позволяют однозначно определить границы посещения.');
    for (const message of fact.messages) warn(message);
    return finish();
  }

  const useSummary = factory || gateFallback === 'summary';
  const covered = !gapDays.length && (useSummary
    ? checkStart != null && sourceRange && sourceRange.start <= checkStart && sourceRange.end >= checkEnd && sourceDates.every(d => summaryRecord.days[d]?.length) && !issues.length
    : covers(checkStart, checkEnd));
  const explainCoverage = () => coverageNote(checkStart, checkEnd, sourceRange, date, gapDays);

  if (!fact.observed) {
    const positive = summaryRows.find(r => r.hasPresence);
    const gateMarked = !factory && person.matches.some(p => p.events.some(e => checkStart != null && e.ts >= checkStart - 3 * HOUR && e.ts <= checkEnd + 3 * HOUR));
    const eligible = (person.events.length || person.summary?.days[date]?.length) && !person.matchConflict && !person.ambiguousSummary && !issues.length && (!plan.conflict || plan.hoursConflict);
    const missing = plan.isWork && covered && eligible && !positive && !gateMarked;
    if (plan.isWork) {
      const window = plan.monthlyWindow || (plan.start != null ? plan : null);
      const planned = `На ${displayDate(date)} назначена смена${window ? ` ${windowText(window, date)}` : ''}.`;
      comments.push(`${planned} В доступной части ${factory ? 'сводки PERCo' : 'выгрузки подробных проходов'} нет отметок входа и выхода за эту смену.`);
    }
    if (person.factSource === 'Только месячный график') comments.push('Сотрудник есть в месячном графике, но не найден ни в одном из файлов проходов.');
    if (positive) comments.push(`В сводке за ${displayDate(positive.date)} указано время присутствия ${positive.presence}, но нет отметок, по которым можно проверить вход и выход.`);
    if (gateMarked) comments.push('В сводке есть отметки проходной рядом с этой сменой. По ним нельзя подтвердить вход в цех.');
    if (plan.isWork && !covered && person.factSource !== 'Только месячный график' && !issues.length) comments.push(explainCoverage());
    if (!missing) needsReview = true;
    status = missing ? 'Не вышел на смену' : 'Недостаточно данных';
    level = missing ? 'red' : 'gray';
    return finish();
  }

  if (plan.knownOff) {
    const reason = day?.code === 'в' ? 'выходной' : day?.code === 'от' ? 'отпуск' : day?.code === '0' ? '0 рабочих часов' : 'на другое время назначена плановая смена';
    deviation('Выход вне графика', `В месячном графике на ${displayDate(date)} ${reason}. В выгрузке есть отметки о проходе.`);
  } else if (!plan.isWork && !plan.normOnly) {
    status = 'Нет нормы для сравнения';
    if (!person.plan) {
      if (!person.matchConflict) level = 'gray';
      warn('Сотрудник не найден в месячном графике. Подходящей нормы в столбце K тоже нет.');
    }
    else if (!plan.conflict) warn(`В месячном графике на ${displayDate(date)} пустая ячейка. Уточни, была ли назначена смена.`);
  } else if (plan.start == null) {
    status = 'Нет нормы для сравнения';
  } else if (plan.normOnly) {
    warn('Месячного графика на эту дату нет. По K можно проверить часы, но неизвестно, была ли назначена смена.');
  }

  if ((plan.isWork || plan.checksTime) && plan.start != null && !person.matchConflict && !person.ambiguousSummary && !issues.length) {
    const wrongShift = fact.first && fact.startGood && !!plan.night !== !!fact.night;
    if (wrongShift) {
      deviation('Другая смена', `По графику ${windowText(plan, date)}, а ${entry} отмечен в ${moment(fact.first.ts, date)}${fact.last ? ` и ${exit} в ${moment(fact.last.ts, date)}` : ''}.`);
    } else {
      const late = fact.first && fact.startGood ? Math.floor(fact.first.ts / MINUTE) - Math.floor(plan.start / MINUTE) : 0;
      const early = fact.last ? Math.floor(plan.end / MINUTE) - Math.floor(fact.last.ts / MINUTE) : 0;
      if (late > 0) deviation('Поздний вход', `Поздний ${entry} на ${late} мин. По графику ${moment(plan.start, date)}, по отметке ${moment(fact.first.ts, date)}.`);
      if (early > 0 && covered && fact.boundaryGood) deviation(late > 0 ? 'Поздний вход и ранний выход' : 'Ранний выход', `Ранний ${exit} на ${early} мин. По графику ${moment(plan.end, date)}, по отметке ${moment(fact.last.ts, date)}.`);
      if (early > 0 && !covered && fact.boundaryGood) warn(`Последний ${exit} в ${moment(fact.last.ts, date)}, по графику до ${moment(plan.end, date)}. По неполной выгрузке ранний уход подтвердить нельзя.`);
    }
  }
  if (!fact.first || !fact.last || !fact.boundaryGood || fact.messages.length || issues.length || ((plan.isWork || plan.checksTime) && !covered)) {
    if (level !== 'red' && status !== 'Нет нормы для сравнения') status = 'Недостаточно данных';
    if (!fact.first) warn(`В выгрузке нет отметки «${entry}» за эту смену.`);
    if (!fact.last) warn(`В выгрузке нет отметки «${exit}» за эту смену.`);
    if (!factory && (!fact.first || !fact.last) && fact.all.some(e => e.from === OUTSIDE || e.to === OUTSIDE)) {
      warn('Отметки проходной подтверждают присутствие на заводе. Для расчёта времени в цехе нужна вторая цеховая отметка.');
    }
    if (fact.first && fact.last && !fact.boundaryGood) warn('Последовательность отметок входа и выхода не позволяет однозначно определить границы смены.');
    if ((plan.isWork || plan.checksTime) && !covered && plan.start != null && !issues.length) warn(explainCoverage());
    for (const message of fact.messages) warn(message);
  }
  return finish();
}

export function buildShiftReport(events, schedule, presence, fromDate, toDate) {
  if (!fromDate || !toDate || fromDate > toDate) throw new Error('Проверь даты отчётного периода.');
  const sorted = dedupeEvents(events), people = reportPeople(sorted, schedule, presence), rows = [];
  const detailedStart = sorted[0]?.ts ?? Infinity, detailedEnd = sorted.at(-1)?.ts ?? -Infinity;
  const range = reportCoverage(events, presence), coverageStart = range?.start ?? Infinity, coverageEnd = range?.end ?? -Infinity;
  // A monthly plan is not evidence for dates beyond the uploaded passes.
  if (!range) return { rows, shops: [], coverageStart, coverageEnd, warnings: schedule?.warnings || [] };
  const throughDate = localDate(coverageEnd);
  if (toDate > throughDate) toDate = throughDate;
  const coverageRanges = [...new Map(events.filter(e=>e.fileCoverage).map(e=>[`${e.fileCoverage.start}|${e.fileCoverage.end}`,e.fileCoverage])).values()].sort((a,b)=>a.start-b.start);
  const detailedDays = new Set(sorted.map(e => localDate(e.ts)));
  const summaryDays = new Set((presence?.people || []).flatMap(person => person.events.map(e => localDate(e.ts))));
  const covers = (start,end) => {
    if(start==null)return false;
    if(!coverageRanges.length)return detailedStart<=start&&detailedEnd>=end;
    let cursor=start;
    for(const range of coverageRanges){if(range.end<cursor)continue;if(range.start>cursor)return false;cursor=Math.max(cursor,range.end);if(cursor>=end)return true;}
    return false;
  };
  for (const person of people) {
    const accounting = person.summary ? 'factory' : 'shop';
    const sourceRange = person.summary ? presence.coverage : { start:detailedStart, end:detailedEnd };
    const sourceEnd = person.factSource === 'Только месячный график' ? coverageEnd : sourceRange?.end ?? (person.plan ? coverageEnd : null);
    const matchedSummary = person.matches.length === 1 && !person.matchConflict && !person.ambiguousSummary ? person.matches[0] : null;
    const summaryRecord = person.summary || matchedSummary;
    if (!Number.isFinite(sourceEnd) && !summaryRecord) continue;
    const profiles = new Map();
    const profileAt = date => {
      if (!profiles.has(date)) profiles.set(date, profileFor(person.matches, date));
      return profiles.get(date);
    };
    const planDays = person.plan?.days || {};
    const primaryGroups = shiftGroups(person.events, profileAt, planDays, fromDate, Boolean(person.summary));
    const summaryGroups = !person.summary && matchedSummary
      ? shiftGroups(matchedSummary.events, profileAt, planDays, fromDate, true)
      : new Map();
    const markDates = new Set((summaryRecord?.events || []).map(e => localDate(e.ts)));
    const recordedDays = new Set(Object.entries(summaryRecord?.days || {})
      .filter(([date, records]) => !markDates.has(date) && records.some(r => r.issues.length || r.hasPresence))
      .map(([date]) => date));
    const dates = new Set([...Object.keys(planDays), ...primaryGroups.keys(), ...summaryGroups.keys(), ...recordedDays]);
    for (const date of [...dates].sort()) {
      if (date < fromDate || date > toDate) continue;
      const detail = primaryGroups.get(date) || [];
      const gate = summaryGroups.get(date);
      const summaryFallback = Boolean(gate?.length && !detail.some(hasShopMarks));
      if (!summaryFallback && !recordedDays.has(date) && (!Number.isFinite(sourceEnd) || date > localDate(sourceEnd))) continue;
      const day = planDays[date], profile = profileAt(date);
      const groups = summaryFallback ? gate : detail;
      const planned = expected(date, day, profile, groups[0]?.[0].night);
      if (!planned.isWork && !planned.unknownCode && !groups.length && !recordedDays.has(date)) continue;
      const jobs = [];
      if (!groups.length) jobs.push({ plan: planned, actual: [] });
      else if (planned.isWork && groups.length > 1) {
        const matched = groups.find(es => es[0].night === planned.night);
        if (matched) jobs.push({ plan: planned, actual: matched });
        else jobs.push({ plan: planned, actual: [] });
        for (const actual of groups) if (actual !== matched) jobs.push({ plan: { isWork: false, knownOff: true, label: 'Вне плановой смены', start: null, end: null }, actual });
      } else for (const actual of groups) jobs.push({ plan: planForActual(date, day, profile, planned, actual), actual });
      for (const { plan, actual } of jobs) {
        const summaryFact = Boolean(person.summary || summaryFallback);
        const detailGate = !summaryFact && actual.some(ep => ep.gateIn || ep.gateOut) && !hasShopMarks(actual);
        let gateFallback = summaryFallback ? 'summary' : detailGate ? 'detail' : null;
        let fact = summaryFact ? factoryFactOf(actual, date) : factOf(actual, date, detailGate);
        if (summaryFallback && actual.length && (!fact.boundaryGood || fact.messages.length)) {
          const detailGroup = detail.find(group => group[0].start === actual[0].start);
          if (detailGroup) {
            const candidate = factOf(detailGroup, date, true);
            if (candidate.boundaryGood && !candidate.messages.length) {
              fact = candidate;
              gateFallback = 'detail';
            }
          }
        }
        if (gateFallback && (!fact.boundaryGood || fact.messages.length)) fact.span = null;
        const result = attendanceResult({ person, plan, day, date, fact, accounting,
          sourceRange:summaryFact ? presence.coverage : sourceRange,
          sourceDays:summaryFact ? summaryDays : detailedDays, covers, gateFallback, summaryRecord });
        const factSource = gateFallback === 'summary' ? 'Сводка PERCo · проходная'
          : gateFallback === 'detail' ? 'Подробные проходы · проходная' : person.factSource;
        const worked = fact.boundaryGood && !fact.messages.length && fact.span != null ? Math.max(0, fact.span - HOUR) : null;
        rows.push({ id:person.id, name:person.name, employeeKey:person.key, department:person.department || '', period:date, date,
          ts:actual[0]?.start ?? plan.start ?? onDate(date), plan, day, profile, accounting,
          factSource, ...fact, worked, ...result,
          planSource:day?.source || '', sources:fact.all.map(e => ref(e.fileName, e.rowNumber, e.sheetName)) });
      }
    }
  }
  return { rows, shops: [...new Set(rows.flatMap(r => Object.keys(r.byShop)))].sort(), coverageStart, coverageEnd, warnings: schedule?.warnings || [] };
}

export function aggregateShops(report, granularity) {
  const groups = new Map();
  for (const row of report.rows) {
    if (row.accounting === 'factory') continue;
    const period = periodKey(row.ts, granularity), key = `${row.employeeKey}|${period}`;
    if (!groups.has(key)) groups.set(key, { id: row.id, name: row.name, employeeKey: row.employeeKey, period, byShop: {}, shopGross: 0, lunchShop: 0, canteen: 0, holding: 0, other: 0, unknown: 0 });
    const g = groups.get(key);
    for (const field of ['shopGross','lunchShop','canteen','holding','other','unknown']) g[field] += row[field];
    for (const [shop, ms] of Object.entries(row.byShop)) g.byShop[shop] = (g.byShop[shop] || 0) + ms;
  }
  return { rows: [...groups.values()], shops: report.shops };
}
