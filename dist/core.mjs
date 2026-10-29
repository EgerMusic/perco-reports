export const OUTSIDE = 'Неконтролируемая территория';
export const HOLDING = 'Территория холдинга';

export function parseStamp(value) {
  if (value instanceof Date) return value.getTime();
  const match = String(value ?? '').trim().match(/^(\d{2})[-.](\d{2})[-.](\d{4})[ T](\d{2}):(\d{2})(?::(\d{2}))?$/);
  if (!match) return NaN;
  const [, d, m, y, h, min, sec = '0'] = match;
  const date = new Date(+y, +m - 1, +d, +h, +min, +sec);
  return date.getFullYear() === +y && date.getMonth() === +m - 1 && date.getDate() === +d ? date.getTime() : NaN;
}

export function parseEvents(rows, fileName = '') {
  const headerIndex = rows.findIndex(row => row.includes('Табельный номер') && row.includes('Выход из') && row.includes('Вход в'));
  if (headerIndex < 0) throw new Error('Не найдены заголовки PERCo: «Табельный номер», «Выход из», «Вход в».');
  const names = rows[headerIndex].map(v => String(v ?? '').trim());
  const col = key => names.indexOf(key);
  const required = ['Дата', 'Табельный номер', 'Выход из', 'Вход в'];
  if (required.some(key => col(key) < 0)) throw new Error('Выгрузка не содержит нужных столбцов PERCo.');
  const events = [], rejected = [];
  for (let i = headerIndex + 1; i < rows.length; i++) {
    const row = rows[i];
    if (row.every(v => v == null || String(v).trim() === '')) continue;
    const ts = parseStamp(row[col('Дата')]);
    const id = String(row[col('Табельный номер')] ?? '').trim();
    const from = String(row[col('Выход из')] ?? '').trim();
    const to = String(row[col('Вход в')] ?? '').trim();
    if (!Number.isFinite(ts) || !id || !from || !to) { rejected.push(i + 1); continue; }
    const name = ['Фамилия', 'Имя', 'Отчество'].map(key => String(row[col(key)] ?? '').trim()).filter(Boolean).join(' ');
    const card = String(row[col('Карта №')] ?? '').trim();
    events.push({ id, name, card, ts, from, to, fileName, rowNumber: i + 1, key: [id, ts, from, to, card].join('|') });
  }
  return { events, rejected };
}

export function dedupeEvents(events) {
  const map = new Map();
  for (const event of events) if (!map.has(event.key)) map.set(event.key, event);
  return [...map.values()].sort((a, b) => a.ts - b.ts || a.id.localeCompare(b.id));
}

function normalized(zone) { return zone === 'Гостевой' ? HOLDING : zone; }
function zoneKind(zone) {
  if (zone === OUTSIDE) return 'outside';
  if (zone === HOLDING || zone === 'Гостевой') return 'holding';
  if (/столов/i.test(zone)) return 'canteen';
  if (/^цех(?:\s|$)/i.test(zone)) return 'shop';
  return 'other';
}
export function issueKey(a, b) { return `${a.id}|${a.ts}|${b.ts}|${a.to}|${b.from}`; }

export function analyze(events, corrections = {}) {
  const byEmployee = new Map();
  for (const event of dedupeEvents(events)) {
    if (!byEmployee.has(event.id)) byEmployee.set(event.id, []);
    byEmployee.get(event.id).push(event);
  }
  const sessions = [], segments = [], issues = [];
  for (const [id, list] of byEmployee) {
    let startIndex = -1;
    for (let i = 0; i < list.length; i++) {
      const event = list[i];
      const entry = event.from === OUTSIDE && event.to !== OUTSIDE;
      const exit = event.from !== OUTSIDE && event.to === OUTSIDE;
      if (entry) {
        if (startIndex >= 0) issues.push({ type: 'repeat-entry', id, name: event.name, a: list[startIndex], b: event, key: issueKey(list[startIndex], event) });
        startIndex = i;
      } else if (exit) {
        if (startIndex < 0) {
          issues.push({ type: 'orphan-exit', id, name: event.name, a: event, b: null, key: `${id}|orphan|${event.ts}` });
          continue;
        }
        const entryEvent = list[startIndex];
        if (event.ts > entryEvent.ts) {
          const session = { id, name: event.name || entryEvent.name, start: entryEvent.ts, end: event.ts, entry: entryEvent, exit: event };
          sessions.push(session);
          for (let j = startIndex; j < i; j++) {
            const a = list[j], b = list[j + 1];
            if (b.ts <= a.ts) continue;
            const from = normalized(a.to), to = normalized(b.from);
            if (from === to) {
              segments.push({ id, name: session.name, start: a.ts, end: b.ts, zone: from, kind: zoneKind(from), sourceStart: a, sourceEnd: b });
            } else {
              const key = issueKey(a, b);
              const cut = Number(corrections[key]?.cut ?? corrections[key]);
              if (Number.isFinite(cut) && cut > a.ts && cut < b.ts && from !== OUTSIDE && to !== OUTSIDE) {
                segments.push({ id, name: session.name, start: a.ts, end: cut, zone: from, kind: zoneKind(from), sourceStart: a, sourceEnd: b, corrected: true });
                segments.push({ id, name: session.name, start: cut, end: b.ts, zone: to, kind: zoneKind(to), sourceStart: a, sourceEnd: b, corrected: true });
              } else {
                segments.push({ id, name: session.name, start: a.ts, end: b.ts, zone: 'Не распределено', kind: 'unknown', sourceStart: a, sourceEnd: b });
                issues.push({ type: 'zone-gap', id, name: session.name, a, b, key });
              }
            }
          }
        }
        startIndex = -1;
      }
    }
    if (startIndex >= 0) {
      const event = list[startIndex];
      issues.push({ type: 'open-entry', id, name: event.name, a: event, b: null, key: `${id}|open|${event.ts}` });
    }
  }
  return { sessions, segments, issues };
}

export function localDate(ts) {
  const d = new Date(ts);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}
export function periodKey(ts, granularity) {
  const d = new Date(ts);
  if (granularity === 'month') return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
  if (granularity === 'week') {
    const day = (d.getDay() + 6) % 7;
    d.setDate(d.getDate() - day);
    return `Неделя с ${localDate(d.getTime())}`;
  }
  return localDate(ts);
}
function slices(start, end, emit) {
  let current = start;
  while (current < end) {
    const d = new Date(current);
    const next = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1).getTime();
    const stop = Math.min(next, end);
    emit(current, stop);
    current = stop;
  }
}

export function buildReport(analysis, startDate, endDate, granularity = 'day') {
  const start = new Date(`${startDate}T00:00:00`).getTime();
  const end = new Date(new Date(`${endDate}T00:00:00`).getFullYear(), new Date(`${endDate}T00:00:00`).getMonth(), new Date(`${endDate}T00:00:00`).getDate() + 1).getTime();
  if (!Number.isFinite(start) || !Number.isFinite(end) || end <= start) throw new Error('Проверь даты отчётного периода.');
  const groups = new Map(), shops = new Set();
  const get = (id, name, ts) => {
    const period = periodKey(ts, granularity), key = `${id}|${period}`;
    if (!groups.has(key)) groups.set(key, { id, name, period, plant: 0, holding: 0, canteen: 0, unknown: 0, other: 0, shopGross: 0, lunchShop: 0, byShop: {}, shifts: 0, visits: [] });
    return groups.get(key);
  };
  for (const session of analysis.sessions) {
    if (session.end > start && session.start < end) slices(Math.max(start, session.start), Math.min(end, session.end), (a, b) => {
      if (b <= a) return;
      const row = get(session.id, session.name, a);
      row.plant += b - a;
      if (!row.visits.some(v => v.start === session.start && v.end === session.end)) row.visits.push(session);
    });
    if (session.start >= start && session.start < end) get(session.id, session.name, session.start).shifts++;
    const parts = analysis.segments.filter(s => s.id === session.id && s.start >= session.start && s.end <= session.end);
    const canteen = parts.filter(s => s.kind === 'canteen').reduce((n, s) => n + s.end - s.start, 0);
    const shop = parts.filter(s => s.kind === 'shop').reduce((n, s) => n + s.end - s.start, 0);
    const lunchShop = Math.min(shop, Math.max(0, 3600000 - Math.min(canteen, 3600000)));
    const shopDays = new Map();
    for (const part of parts.filter(s => s.kind === 'shop')) slices(part.start, part.end, (a, b) => {
      const day = localDate(a);
      shopDays.set(day, (shopDays.get(day) || 0) + b - a);
    });
    const lunchDay = [...shopDays].sort((a, b) => b[1] - a[1])[0]?.[0];
    if (lunchDay) {
      const lunchTime = new Date(`${lunchDay}T12:00:00`).getTime();
      if (lunchTime >= start && lunchTime < end) get(session.id, session.name, lunchTime).lunchShop += lunchShop;
    }
  }
  for (const segment of analysis.segments) {
    if (segment.kind === 'shop') shops.add(segment.zone);
    if (segment.end <= start || segment.start >= end) continue;
    slices(Math.max(start, segment.start), Math.min(end, segment.end), (a, b) => {
      const item = get(segment.id, segment.name, a), ms = b - a;
      if (segment.kind === 'shop') { item.shopGross += ms; item.byShop[segment.zone] = (item.byShop[segment.zone] || 0) + ms; }
      else if (segment.kind === 'canteen') item.canteen += ms;
      else if (segment.kind === 'holding') item.holding += ms;
      else if (segment.kind === 'unknown') item.unknown += ms;
      else item.other += ms;
    });
  }
  return { rows: [...groups.values()].sort((a, b) => a.period.localeCompare(b.period) || a.name.localeCompare(b.name)), shops: [...shops].sort() };
}

export function duration(ms) {
  const minutes = Math.round(Math.max(0, ms) / 60000);
  return `${Math.floor(minutes / 60)}:${String(minutes % 60).padStart(2, '0')}`;
}
