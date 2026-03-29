const MONTHS_DE = ['Januar','Februar','März','April','Mai','Juni','Juli','August','September','Oktober','November','Dezember'];
const DAYS_SHORT = ['So','Mo','Di','Mi','Do','Fr','Sa'];

let curYear, curMonth;
let currentState = 'DE-BW';
let holidaysCache = {};
let absences = [];
let viewMode = 'month';
const MAX_FUTURE_DAYS = 180;
let settings = {
  hoursPerWeek: 40,
  workDays: [1,2,3,4,5],
  customHolidays: {},
  colors: {
    accent: '#2563eb',
    accent2: '#2563eb',
    vacation: '#004D40',
    'vacation-bg': '#80CBC4',
    'vacation-border': '#4DB6AC',
    sick: '#880E4F',
    'sick-bg': '#F48FB1',
    'sick-border': '#F06292',
    holiday: '#BF360C',
    'holiday-bg': '#FFF3E0',
    'holiday-border': '#FFB74D',
    office: '#1A237E',
    'office-bg': '#C5CAE9',
    'office-border': '#7986CB',
    other: '#7c3aed',
    'other-bg': '#ede9fe',
    'other-border': '#c4b5fd',
    note: '#713f12',
    'note-bg': '#fef9c3',
    'note-border': '#fcd34d',
    bg: '#eef4fd'
  }
};

const DEFAULT_COLORS = JSON.parse(JSON.stringify(settings.colors));

const pad = n => String(n).padStart(2,'0');
const mkDate = (y,m,d) => `${y}-${pad(m+1)}-${pad(d)}`;

// ── Vollzeit-Sonderregel: ab 35h immer 8 Tage, darunter proportional ────────
function calcBaseDays() {
  return settings.hoursPerWeek >= 35 ? 8 : Math.round(settings.hoursPerWeek / 5);
}

function saveData() {
  localStorage.setItem('officeTracker', JSON.stringify({
    absences,
    view: { year: curYear, month: curMonth, state: currentState, mode: viewMode },
    settings
  }));
}

function loadData() {
  const saved = localStorage.getItem('officeTracker');
  if (saved) {
    try {
      const data = JSON.parse(saved);
      if (data.absences !== undefined) {
        try { validateImport({ absences: data.absences }); } catch (e) {
          console.warn('Gespeicherte Abwesenheiten ungültig, werden zurückgesetzt:', e);
          data.absences = [];
        }
      }
      absences = data.absences || [];
      if (data.view) {
        curYear = typeof data.view.year === 'number' && data.view.year >= 2000 && data.view.year <= 2100 ? data.view.year : new Date().getFullYear();
        curMonth = typeof data.view.month === 'number' && data.view.month >= 0 && data.view.month <= 11 ? data.view.month : new Date().getMonth();
        currentState = data.view.state || 'DE-BW';
        viewMode = data.view.mode || 'month';
      }
      if (data.settings) {
        settings = { ...settings, ...data.settings };
      }
    } catch (e) {
      console.error('Fehler beim Laden der Daten:', e);
    }
  }
  applyCustomColors();
}

function applyCustomColors() {
  const root = document.documentElement;
  root.style.setProperty('--accent', settings.colors.accent);
  root.style.setProperty('--accent2', settings.colors.accent2 || settings.colors.accent);
  root.style.setProperty('--vacation', settings.colors.vacation);
  root.style.setProperty('--vacation-bg', settings.colors['vacation-bg']);
  root.style.setProperty('--vacation-border', settings.colors['vacation-border']);
  root.style.setProperty('--sick', settings.colors.sick);
  root.style.setProperty('--sick-bg', settings.colors['sick-bg']);
  root.style.setProperty('--sick-border', settings.colors['sick-border']);
  root.style.setProperty('--holiday', settings.colors.holiday);
  root.style.setProperty('--holiday-bg', settings.colors['holiday-bg']);
  root.style.setProperty('--holiday-border', settings.colors['holiday-border']);
  root.style.setProperty('--office', settings.colors.office);
  root.style.setProperty('--office-bg', settings.colors['office-bg']);
  root.style.setProperty('--office-border', settings.colors['office-border']);
  root.style.setProperty('--other', settings.colors.other || '#7c3aed');
  root.style.setProperty('--other-bg', settings.colors['other-bg'] || '#ede9fe');
  root.style.setProperty('--other-border', settings.colors['other-border'] || '#c4b5fd');
  root.style.setProperty('--note', settings.colors.note || '#713f12');
  root.style.setProperty('--note-bg', settings.colors['note-bg'] || '#fef9c3');
  root.style.setProperty('--note-border', settings.colors['note-border'] || '#fcd34d');
  root.style.setProperty('--bg', settings.colors.bg);
}

function isWorkday(dow) {
  return settings.workDays.includes(dow);
}

async function ensureHolidays(year) {
  if (holidaysCache[year]) return holidaysCache[year];
  const h = {};
  Object.entries(settings.customHolidays).forEach(([date, name]) => {
    if (date.startsWith(String(year))) h[date] = name;
  });
  try {
    const subdivisionCode = currentState;
    const url = `https://openholidaysapi.org/PublicHolidays?countryIsoCode=DE&subdivisionCode=${encodeURIComponent(subdivisionCode)}&languageIsoCode=DE&validFrom=${year}-01-01&validTo=${year}-12-31`;
    const r = await fetch(url);
    if (r.ok) {
      const data = await r.json();
      data.forEach(entry => {
        const date = entry.startDate;
        const name = entry.name?.find(n => n.language === 'DE')?.text
                  || entry.name?.[0]?.text
                  || 'Feiertag';
        h[date] = name;
      });
    }
  } catch (e) {
    console.warn('Feiertage konnten nicht geladen werden:', e);
  }
  holidaysCache[year] = h;
  return h;
}

async function loadAndRender() {
  const now = new Date();

  if (viewMode === 'year') {
    document.getElementById('cal-month-label').textContent = `${curYear}`;
    document.getElementById('cal-header-left').textContent = 'Jahresansicht';
    document.getElementById('today-btn').style.display = curYear === now.getFullYear() ? 'none' : 'inline-flex';
  } else {
    document.getElementById('cal-month-label').textContent = `${MONTHS_DE[curMonth]} ${curYear}`;
    document.getElementById('cal-header-left').textContent = 'Monatsansicht';
    const isCurrent = curMonth === now.getMonth() && curYear === now.getFullYear();
    document.getElementById('today-btn').style.display = isCurrent ? 'none' : 'inline-flex';
  }

  await ensureHolidays(curYear);

  if (viewMode === 'year') {
    await renderYearOverview();
  } else {
    renderMonth();
  }
  updateStats();

  document.getElementById('view-toggle').textContent = viewMode === 'month' ? 'Jahr' : 'Monat';
}

function isHoliday(dateStr) {
  const year = dateStr.split('-')[0];
  return !!holidaysCache[year]?.[dateStr];
}

function addDropListeners(el, dateStr, noteOnly = false) {
  el.addEventListener('dragover', e => { e.preventDefault(); el.closest('.day-row').classList.add('drag-over'); });
  el.addEventListener('dragleave', () => el.closest('.day-row').classList.remove('drag-over'));
  el.addEventListener('drop', e => {
    e.preventDefault();
    el.closest('.day-row').classList.remove('drag-over');
    const type = e.dataTransfer.getData('type');
    if (type === 'note') addNote(dateStr);
    else if (!noteOnly && type) addEntry(type, dateStr);
  });
}

function renderMonth() {
  const body = document.getElementById('cal-body');
  const now = new Date();
  const todayStr = mkDate(now.getFullYear(), now.getMonth(), now.getDate());
  const dim = new Date(curYear, curMonth+1, 0).getDate();
  let lastKW = null;
  body.innerHTML = '';

  for(let d=1; d<=dim; d++){
    const date = new Date(curYear, curMonth, d);
    const dow = date.getDay();
    const isWE = dow===0 || dow===6;
    const isWork = isWorkday(dow);
    const dateStr = mkDate(curYear, curMonth, d);
    const isToday = dateStr === todayStr;
    const hol = !!holidaysCache[curYear]?.[dateStr];
    const absInfo = getAbsInfo(dateStr);
    const noteInfo = getNoteInfo(dateStr);
    const isOffice = absInfo && absInfo.ab.type === 'office';
    const isOther  = absInfo && absInfo.ab.type === 'other';
    const isVacation = absInfo && absInfo.ab.type === 'vacation';
    const isSick = absInfo && absInfo.ab.type === 'sick';
    const kw = getWeekNum(date);
    const showKW = isWork && kw !== lastKW;
    if(isWork) lastKW = kw;

    const hasNote = !!noteInfo;
    const isSplitDay = hasNote && (hol || !!absInfo);

    const row = document.createElement('div');
    row.className = 'day-row'
      + (isWE ? ' weekend' : '')
      + (isToday ? ' is-today' : '')
      + (!hasNote && isOffice ? ' office-day' : '')
      + (!hasNote && isVacation ? ' vacation-day' : '')
      + (!hasNote && isSick ? ' sick-day' : '')
      + (!hasNote && isOther ? ' other-day' : '')
      + (isSplitDay && isOffice ? ' office-day split-note' : '')
      + (isSplitDay && isVacation ? ' vacation-day split-note' : '')
      + (isSplitDay && isSick ? ' sick-day split-note' : '')
      + (isSplitDay && isOther ? ' other-day split-note' : '')
      + (isSplitDay && hol ? ' split-note' : '')
      + (!isSplitDay && hasNote ? ' note-day' : '')
      + (isWork && !hol && !absInfo && !hasNote ? ' droppable' : '');
    row.dataset.date = dateStr;

    const kwCell = document.createElement('div');
    kwCell.className = 'kw-cell';
    kwCell.textContent = showKW ? `${kw}` : '';

    const dayCell = document.createElement('div');
    dayCell.className = 'day-cell';
    dayCell.innerHTML = `<span class="day-num-big">${d}</span><span class="day-name-small">${DAYS_SHORT[dow]}</span>${isToday ? '<div class="today-badge"></div>' : ''}`;

    const absCell = document.createElement('div');
    absCell.className = 'absence-cell';

    if (hol) {
      const bl = document.createElement('div');
      bl.className = 'ab-block holiday';
      bl.innerHTML = `<span class="ab-label">🎉 ${escapeHtml(holidaysCache[curYear][dateStr])}</span>`;
      if (hasNote) {
        absCell.classList.add('has-note');
        absCell.appendChild(bl);
        absCell.appendChild(buildNoteBlock(noteInfo, true));
      } else {
        absCell.appendChild(bl);
        addDropListeners(absCell, dateStr, true);
      }
    } else if (absInfo) {
      const {ab, idx} = absInfo;
      const total = ab.dates.length;
      const bl = document.createElement('div');
      bl.className = `ab-block ${ab.type}`;
      const { emoji, label: baseLabel } = TYPE_META[ab.type] || { emoji: '📋', label: 'Sonstige Abw.' };
      const label = ab.type === 'other' ? (ab.customLabel || 'Sonstige Abw.') : baseLabel;
      const labelHtml = ab.type === 'other'
        ? idx === 0
          ? `<input class="ab-label-input" data-id="${ab.id}" value="${(ab.customLabel||'').replace(/"/g,'&quot;')}" placeholder="Sonstige Abw." title="Bezeichnung ändern">`
          : `<span class="ab-label" style="opacity:0.6;font-size:10px;">${escapeHtml(ab.customLabel || 'Sonstige')}</span>`
        : `<span class="ab-label">${emoji} ${escapeHtml(label)}</span>`;

      // Gear button only on first day of block
      const gearHtml = idx === 0
        ? `<button class="gear-btn" data-id="${ab.id}" data-date="${dateStr}" data-type="${ab.type}" title="Serie erstellen">${GEAR_SVG}</button>`
        : '';

      bl.innerHTML = `
        ${labelHtml}
        ${idx === 0 ? `
        <div class="ab-controls">
          ${gearHtml}
          <button class="ctrl-btn btn-minus" data-id="${ab.id}">−</button>
          <span class="ab-day-count">${total}d</span>
          <button class="ctrl-btn btn-plus" data-id="${ab.id}">+</button>
          <button class="ctrl-del btn-del" data-id="${ab.id}" style="background:none;border:none;cursor:pointer;display:flex;align-items:center;justify-content:center;padding:0;"><svg width="10" height="10" viewBox="0 0 14 14" fill="none"><path d="M1 1L13 13M1 13L13 1" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg></button>
        </div>` : ''}`;
      if (hasNote) {
        absCell.classList.add('has-note');
        absCell.appendChild(bl);
        absCell.appendChild(buildNoteBlock(noteInfo, true));
      } else {
        absCell.appendChild(bl);
        addDropListeners(absCell, dateStr, true);
      }
    } else if (hasNote) {
      absCell.classList.add('note-only');
      absCell.appendChild(buildNoteBlock(noteInfo, false));
    }

    row.appendChild(kwCell);
    row.appendChild(dayCell);
    row.appendChild(absCell);
    body.appendChild(row);

    if (isWork && !hol && !absInfo) addDropListeners(row, dateStr, false);
    if (isWE && !hasNote) addDropListeners(row, dateStr, true);
  }
}

async function renderYearOverview() {
  const body = document.getElementById('cal-body');
  body.innerHTML = '';
  const grid = document.createElement('div');
  grid.className = 'year-grid';

  const now = new Date();
  const todayMonth = now.getMonth();
  const todayYear = now.getFullYear();
  const baseDays = calcBaseDays();

  for (let m = 0; m < 12; m++) {
    const prefix = `${curYear}-${pad(m+1)}-`;
    const isCurrent = curYear === todayYear && m === todayMonth;
    const isPast    = curYear < todayYear || (curYear === todayYear && m < todayMonth);
    const isFuture  = !isCurrent && !isPast;

    const metrics = getMonthlyMetrics(curYear, m);
    const vacDays   = metrics.counts.vacation;
    const sickDays  = metrics.counts.sick;
    const otherDays = metrics.counts.other;
    const attended  = metrics.counts.office;
    const holidayCount = metrics.holidayCount;
    const pflicht = metrics.officePflicht;

    let status = 'future';
    if (!isFuture) {
      if (attended >= pflicht) status = 'done';
      else if (attended > 0)   status = 'partial';
      else                     status = 'open';
    }
    if (isCurrent && attended >= pflicht) status = 'done';

    const pct = pflicht > 0 ? Math.min(100, Math.round((attended / pflicht) * 100)) : 0;
    const fillClass = attended >= pflicht && pflicht > 0 ? 'done' : attended > 0 ? 'partial' : 'empty';

    let badgeText = '', badgeClass = 'future';
    if (isCurrent)       { badgeText = 'Aktuell'; badgeClass = 'today'; }
    else if (status === 'done')    { badgeText = '✓ Erfüllt'; badgeClass = 'done'; }
    else if (status === 'partial') { badgeText = 'Offen';    badgeClass = 'partial'; }
    else if (isPast && status === 'open') { badgeText = 'Keine Einträge'; badgeClass = 'future'; }

    const stripeClass = isFuture ? 'future' : status === 'done' ? 'done' : status === 'partial' ? 'partial' : 'open';

    const absMap = new Map();
    for (const ab of absences) {
      if (ab.type === 'note') continue;
      for (const date of ab.dates) {
        if (date.startsWith(prefix)) absMap.set(date, ab.type);
      }
    }

    let dotsHtml = '';
    for (let d = 1; d <= dim; d++) {
      const dow = new Date(curYear, m, d).getDay();
      const dStr = mkDate(curYear, m, d);
      const isWE = dow === 0 || dow === 6;
      const isHol = !!holidaysCache[curYear]?.[dStr];
      let dotClass = 'workday';
      if (isWE) dotClass = 'weekend';
      else if (isHol) dotClass = 'holiday';
      else if (absMap.has(dStr)) dotClass = absMap.get(dStr);
      dotsHtml += `<div class="year-dot ${dotClass}" title="${dStr.split('-').reverse().join('.')}"></div>`;
    }

    const numClass = attended >= pflicht && pflicht > 0 ? 'done' : attended > 0 ? 'partial' : '';

    const card = document.createElement('div');
    card.className = 'year-card' + (isCurrent ? ' is-current-month' : '') + (isPast ? ' is-past' : '');
    card.innerHTML = `
      <div class="year-card-stripe ${stripeClass}"></div>
      <div class="year-card-inner">
        <div class="year-card-header">
          <div class="year-card-month">${MONTHS_DE[m]}</div>
          ${badgeText ? `<div class="year-card-badge ${badgeClass}">${badgeText}</div>` : ''}
        </div>
        <div class="year-card-body">
          <div>
            <div class="year-office-big">
              <span class="year-office-num ${numClass}">${attended}</span>
              <span class="year-office-denom"> / ${pflicht}</span>
            </div>
            <div class="year-office-sublabel">🏢 Office Tage</div>
          </div>
          <div class="year-chips">
            ${vacDays  > 0 ? `<div class="year-chip vacation">🌴 ${vacDays}d</div>` : ''}
            ${sickDays > 0 ? `<div class="year-chip sick">😷 ${sickDays}d</div>` : ''}
            ${holidayCount > 0 ? `<div class="year-chip holiday">🎉 ${holidayCount}</div>` : ''}
            ${otherDays > 0 ? `<div class="year-chip other">📋 ${otherDays}d</div>` : ''}
          </div>
        </div>
        <div class="year-progress-track">
          <div class="year-progress-fill ${fillClass}" style="width:${pct}%"></div>
        </div>
        <div class="year-dot-row">${dotsHtml}</div>
      </div>
    `;
    card.addEventListener('click', () => {
      curMonth = m;
      viewMode = 'month';
      loadAndRender();
    });
    grid.appendChild(card);
  }
  body.appendChild(grid);
}

async function addEntry(type, startDate) {
  const dates = await computeDates(startDate, 1);
  if (dates.length) absences.push({ id: 'ab_'+Date.now(), type, dates });
  saveData();
  renderMonth();
  updateStats();
}

function getAbsInfo(dateStr, type = null) {
  for (const ab of absences) {
    if (type !== null) {
      if (ab.type !== type) continue;
    } else {
      if (ab.type === 'note') continue;
    }
    const idx = ab.dates.indexOf(dateStr);
    if (idx !== -1) return { ab, idx };
  }
  return null;
}

function getNoteInfo(dateStr) {
  return getAbsInfo(dateStr, 'note');
}

async function addNote(dateStr) {
  if(getNoteInfo(dateStr)) return;
  absences.push({ id: 'ab_'+Date.now(), type: 'note', dates: [dateStr], customLabel: '' });
  saveData();
  renderMonth();
  updateStats();
}

function deleteNote(id) {
  absences = absences.filter(a => a.id !== id);
  saveData();
  renderMonth();
  updateStats();
}

async function extendNote(id, delta) {
  const ab = absences.find(a => a.id === id);
  if (!ab) return;
  if (delta > 0) {
    let last = ab.dates[ab.dates.length - 1];
    let [y, m, d] = last.split('-').map(Number);
    let cursor = new Date(y, m - 1, d + 1);
    const existingNotes = new Set(absences.filter(a => a.type === 'note' && a.id !== id).flatMap(a => a.dates));
    for (let i = 0; i < MAX_FUTURE_DAYS; i++) {
      const dStr = mkDate(cursor.getFullYear(), cursor.getMonth(), cursor.getDate());
      if (!existingNotes.has(dStr)) {
        ab.dates.push(dStr);
        ab.dates.sort((a, b) => a.localeCompare(b));
        break;
      }
      cursor.setDate(cursor.getDate() + 1);
    }
  } else {
    if (ab.dates.length <= 1) {
      absences = absences.filter(a => a.id !== id);
    } else {
      ab.dates.sort((a, b) => a.localeCompare(b));
      ab.dates.pop();
    }
  }
  saveData();
  renderMonth();
  updateStats();
}

function buildNoteBlock(noteInfo, hasOtherEntry) {
  const {ab, idx} = noteInfo;
  const total = ab.dates.length;
  const bl = document.createElement('div');
  bl.className = 'ab-block note' + (hasOtherEntry ? ' note-split' : '');
  bl.innerHTML = `
    ${idx === 0
      ? `<input class="ab-label-input note-label-input" data-note-id="${ab.id}" value="${(ab.customLabel||'').replace(/"/g,'&quot;')}" placeholder="Notiz…">`
      : `<span class="ab-label" style="opacity:0.6;font-size:10px;">${escapeHtml(ab.customLabel || 'Notiz')}</span>`
    }
    ${idx === 0 ? `
    <div class="ab-controls">
      <button class="gear-btn" data-id="${ab.id}" data-date="${ab.dates[0]}" data-type="note" title="Serie erstellen">${GEAR_SVG}</button>
      <button class="ctrl-btn btn-note-minus" data-note-id="${ab.id}">−</button>
      <span class="ab-day-count">${total}d</span>
      <button class="ctrl-btn btn-note-plus" data-note-id="${ab.id}">+</button>
      <button class="ctrl-del btn-note-del" data-note-id="${ab.id}" style="background:none;border:none;cursor:pointer;display:flex;align-items:center;justify-content:center;padding:0;"><svg width="10" height="10" viewBox="0 0 14 14" fill="none"><path d="M1 1L13 13M1 13L13 1" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg></button>
    </div>` : ''}`;
  return bl;
}

function getWeekNum(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay()||7));
  const y = new Date(Date.UTC(d.getUTCFullYear(),0,1));
  return Math.ceil((((d-y)/86400000)+1)/7);
}

async function computeDates(startDate, numDays) {
  const [sy,sm,sd] = startDate.split('-').map(Number);
  const existing = new Set(absences.filter(a => a.type !== 'note').flatMap(a => a.dates));
  const result = [];
  let cursor = new Date(sy, sm-1, sd);
  let added = 0;

  while(added < numDays) {
    const dow = cursor.getDay();
    const dStr = mkDate(cursor.getFullYear(), cursor.getMonth(), cursor.getDate());
    const y = cursor.getFullYear();

    await ensureHolidays(y);
    const hol = isHoliday(dStr);

    if(isWorkday(dow) && !hol && !existing.has(dStr)){
      result.push(dStr);
      added++;
    }
    cursor.setDate(cursor.getDate() + 1);
  }
  return result;
}

let isExtending = false;

async function extendAbsence(id, delta) {
  if (isExtending) return;
  isExtending = true;
  try {
    const ab = absences.find(a => a.id === id);
    if (!ab) return;

    if (delta > 0) {
      let last = ab.dates[ab.dates.length - 1];
      let [y, m, d] = last.split('-').map(Number);
      let cursor = new Date(y, m - 1, d + 1);

      const existing = new Set(absences.filter(a => a.type !== 'note').flatMap(a => a.dates));

      for (let i = 0; i < MAX_FUTURE_DAYS; i++) {
        const dow = cursor.getDay();
        const dStr = mkDate(cursor.getFullYear(), cursor.getMonth(), cursor.getDate());
        const cy = cursor.getFullYear();

        await ensureHolidays(cy);
        const hol = isHoliday(dStr);

        if (isWorkday(dow) && !hol && !existing.has(dStr)) {
          ab.dates.push(dStr);
          ab.dates.sort((a, b) => a.localeCompare(b));
          break;
        }
        cursor.setDate(cursor.getDate() + 1);
      }
    } else {
      if (ab.dates.length <= 1) {
        absences = absences.filter(a => a.id !== id);
      } else {
        ab.dates.sort((a, b) => a.localeCompare(b));
        ab.dates.pop();
      }
    }

    saveData();
    renderMonth();
    updateStats();
  } finally {
    isExtending = false;
  }
}

function deleteAbsence(id) {
  absences = absences.filter(a => a.id !== id || a.type === 'note');
  saveData();
  renderMonth();
  updateStats();
}

function getMonthlyMetrics(year, month) {
  const dim = new Date(year, month + 1, 0).getDate();
  let workdays = 0, holidayCount = 0;
  const yearHolidays = holidaysCache[year] || {};

  for (let d = 1; d <= dim; d++) {
    const dStr = mkDate(year, month, d);
    if (isWorkday(new Date(year, month, d).getDay())) {
      workdays++;
      if (yearHolidays[dStr]) holidayCount++;
    }
  }

  const prefix = `${year}-${pad(month + 1)}-`;
  const counts = { vacation: 0, sick: 0, other: 0, office: 0 };
  absences.forEach(ab => {
    if (counts.hasOwnProperty(ab.type)) {
      counts[ab.type] += ab.dates.filter(d => d.startsWith(prefix)).length;
    }
  });

  const totalAbs = counts.vacation + counts.sick + counts.other;
  const baseDays = calcBaseDays();
  const effectiveWorkdays = workdays - holidayCount;
  const raw = effectiveWorkdays === 0 ? baseDays : baseDays - (totalAbs * (baseDays / effectiveWorkdays));
  const officePflicht = Math.round(Math.max(0, raw));

  return {
    workdays, holidayCount, counts, totalAbs, officePflicht, baseDays,
    netWorkdays: workdays - holidayCount - totalAbs
  };
}

function updateStats() {
  const metrics = getMonthlyMetrics(curYear, curMonth);

  document.getElementById('stat-workdays').textContent = metrics.workdays;
  document.getElementById('stat-holidays').textContent = metrics.holidayCount;
  document.getElementById('stat-absences').textContent = metrics.totalAbs;
  document.getElementById('stat-net').textContent = Math.max(0, metrics.netWorkdays);
  document.getElementById('stat-basedays').textContent = metrics.baseDays;
  document.getElementById('stat-basedays-sub').textContent = settings.hoursPerWeek >= 35
    ? `${settings.hoursPerWeek}h/Woche (VZ = 8)`
    : `${settings.hoursPerWeek}h / 5 Tage`;

  const officeCard = document.getElementById('stat-office').closest('.stat-card');
  if (officeCard) {
    officeCard.classList.remove('status-ok', 'status-warn');
    if (metrics.officePflicht > 0) {
      officeCard.classList.add(metrics.counts.office >= metrics.officePflicht ? 'status-ok' : 'status-warn');
    }
  }
  document.getElementById('stat-office').innerHTML =
    `<span>${metrics.counts.office}</span> / ${metrics.officePflicht}`;
}

// ─── Helpers ────────────────────────────────────────────────────────────────

function escapeHtml(str) {
  return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

const TYPE_META = {
  vacation: { emoji: '🌴', label: 'Urlaub' },
  sick:     { emoji: '😷', label: 'Krank' },
  office:   { emoji: '🏢', label: 'Im Office' },
  other:    { emoji: '📋', label: 'Sonstige Abw.' },
  note:     { emoji: '📌', label: 'Notiz' },
};

const GEAR_SVG = `<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83-2.83l.06-.06A1.65 1.65 0 0 0 4.68 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 2.83-2.83l.06.06A1.65 1.65 0 0 0 9 4.68a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 2.83l-.06.06A1.65 1.65 0 0 0 19.4 9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z"/></svg>`;

const DATE_RE = /^\d{4}-\d{2}-\d{2}$/;
const VALID_TYPES = ['vacation', 'sick', 'office', 'other', 'note'];
const COLOR_KEYS = ['accent', 'accent2', 'vacation', 'vacation-bg', 'vacation-border', 'sick', 'sick-bg', 'sick-border', 'holiday', 'holiday-bg', 'holiday-border', 'office', 'office-bg', 'office-border', 'other', 'other-bg', 'other-border', 'note', 'note-bg', 'note-border', 'bg'];

function validateImport(data) {
  if (typeof data !== 'object' || data === null) throw new Error('Ungültiges Format');
  if (data.absences !== undefined) {
    if (!Array.isArray(data.absences)) throw new Error('"absences" muss ein Array sein');
    for (const ab of data.absences) {
      if (typeof ab.id !== 'string') throw new Error('Ungültige Abwesenheits-ID');
      if (!VALID_TYPES.includes(ab.type)) throw new Error(`Ungültiger Typ: ${ab.type}`);
      if (!Array.isArray(ab.dates)) throw new Error('Abwesenheits-Daten müssen ein Array sein');
      for (const d of ab.dates) {
        if (!DATE_RE.test(d)) throw new Error(`Ungültiges Datumsformat: ${d}`);
      }
    }
  }
  if (data.view !== undefined && typeof data.view !== 'object') throw new Error('"view" muss ein Objekt sein');
  if (data.settings !== undefined) {
    if (typeof data.settings !== 'object') throw new Error('"settings" muss ein Objekt sein');
    if (data.settings.hoursPerWeek !== undefined && (typeof data.settings.hoursPerWeek !== 'number' || data.settings.hoursPerWeek < 1 || data.settings.hoursPerWeek > 60)) {
      throw new Error('Ungültiger Wert für hoursPerWeek');
    }
  }
}

function renderCustomHolidays() {
  const list = document.getElementById('custom-holidays-list');
  const entries = Object.entries(settings.customHolidays);
  if (!entries.length) { list.innerHTML = ''; }
  else {
    list.innerHTML = entries.map(([date, name]) => `
      <div class="custom-holiday-item">
        <span>🎉 ${escapeHtml(name)} <span class="custom-holiday-date">(${escapeHtml(date)})</span></span>
        <button class="remove-holiday-btn" data-date="${escapeHtml(date)}" style="background:none;border:none;cursor:pointer;display:flex;align-items:center;justify-content:center;padding:0;font-size:14px;color:var(--sick);"><svg width="12" height="12" viewBox="0 0 14 14" fill="none"><path d="M1 1L13 13M1 13L13 1" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg></button>
      </div>
    `).join('');
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// SERIES PANEL
// ═══════════════════════════════════════════════════════════════════════════

let seriesPanel = null;
let seriesOverlay = null;
let seriesPanelState = null;
// seriesPanelState: { id, type, startDate, sourceDates, n, untilYear, untilMonth,
//                     minYear, minMonth, maxYear, maxMonth, insertedIds }

function spMonthLabel(y, m) {
  return `${MONTHS_DE[m]} ${y}`;
}

function closeSeriesPanel() {
  if (seriesOverlay) {
    seriesOverlay.classList.remove('series-overlay-visible');
    setTimeout(() => { seriesOverlay?.remove(); seriesOverlay = null; }, 200);
  }
  if (seriesPanel) {
    seriesPanel.classList.remove('series-panel-visible');
    setTimeout(() => { seriesPanel?.remove(); seriesPanel = null; }, 200);
  }
  document.querySelectorAll('.gear-btn').forEach(b => b.style.opacity = '');
}

async function openSeriesPanel(gearBtn, id, type, startDate) {
  closeSeriesPanel();
  gearBtn.style.opacity = '0';

  // Get the full source block dates
  const sourceAb = absences.find(a => a.id === id);
  const sourceDates = sourceAb ? [...sourceAb.dates].sort() : [startDate];

  // Calculate minimum N (number of weeks) based on block size
  const [fy, fm, fd] = sourceDates[0].split('-').map(Number);
  const [ly, lm, ld] = sourceDates[sourceDates.length - 1].split('-').map(Number);
  const firstKW = getWeekNum(new Date(fy, fm - 1, fd));
  const lastKW = getWeekNum(new Date(ly, lm - 1, ld));
  const minN = lastKW - firstKW + 1;

  // Until bounds: min = current month, max = current month + 24
  const now = new Date();
  const minYear = now.getFullYear();
  const minMonth = now.getMonth();
  const maxDate = new Date(now.getFullYear(), now.getMonth() + 24, 1);
  const maxYear = maxDate.getFullYear();
  const maxMonth = maxDate.getMonth();

  // Default until = December of current year (clamped to max)
  let untilYear = minYear;
  let untilMonth = 11;
  if (untilYear > maxYear || (untilYear === maxYear && untilMonth > maxMonth)) {
    untilYear = maxYear; untilMonth = maxMonth;
  }

  seriesPanelState = {
    id, type,
    startDate: sourceDates[0],
    sourceDates,
    n: Math.max(1, minN),
    minN,
    untilYear, untilMonth,
    minYear, minMonth,
    maxYear, maxMonth,
    insertedIds: null
  };

  // Overlay — blocks all outside interaction
  const overlay = document.createElement('div');
  overlay.className = 'series-overlay';
  document.body.appendChild(overlay);
  seriesOverlay = overlay;
  overlay.addEventListener('click', () => {
    closeSeriesPanel();
    gearBtn.style.opacity = '';
  });

  // Panel
  const panel = document.createElement('div');
  panel.className = 'series-panel';
  panel.innerHTML = buildSeriesPanelHTML();
  document.body.appendChild(panel);
  seriesPanel = panel;

  positionSeriesPanel(panel);

  // Preload holidays for entire range
  const yearsNeeded = new Set();
  for (let y = minYear; y <= maxYear + 1; y++) yearsNeeded.add(y);
  await Promise.all([...yearsNeeded].map(y => ensureHolidays(y)));

  requestAnimationFrame(() => requestAnimationFrame(() => {
    overlay.classList.add('series-overlay-visible');
    panel.classList.add('series-panel-visible');
  }));

  bindSeriesPanelEvents(panel, gearBtn);
  await updateSeriesPreview(panel);
}

function buildSeriesPanelHTML() {
  const { n, sourceDates, untilYear, untilMonth } = seriesPanelState;
  const blockSize = sourceDates.length;
  return `
    <div class="sp-body">
      <div class="sp-title">🔁 Serie erstellen</div>
      ${blockSize > 1 ? `<div class="sp-block-info">Blockgröße: <strong>${blockSize} Tage</strong></div>` : ''}
      <div class="sp-rhythm-row">
        <span class="sp-rhythm-label">Alle</span>
        <button class="sp-n-btn sp-n-minus">−</button>
        <span class="sp-n-val">${n}</span>
        <button class="sp-n-btn sp-n-plus">+</button>
        <span class="sp-rhythm-label">Wochen wiederholen</span>
      </div>
      <div class="sp-until-row">
        <span class="sp-rhythm-label">Bis inkl.</span>
        <div class="sp-month-nav">
          <button class="sp-month-prev nav-btn">←</button>
          <span class="sp-month-label">${spMonthLabel(untilYear, untilMonth)}</span>
          <button class="sp-month-next nav-btn">→</button>
        </div>
      </div>
      <div class="sp-preview">
        <div class="sp-preview-loading">Berechne…</div>
      </div>
      <div class="sp-collisions" style="display:none;"></div>
      <div class="sp-footer">
        <button class="sp-btn-cancel">Abbrechen</button>
        <button class="sp-btn-insert" disabled>Einfügen</button>
      </div>
    </div>
  `;
}

function positionSeriesPanel(panel) {
  const panelW = 340;
  panel.style.width = panelW + 'px';
  panel.style.position = 'fixed';
  panel.style.zIndex = '501';
  panel.style.left = `calc(50% - ${panelW / 2}px)`;
  panel.style.top = `calc(50% - 180px)`;
  panel.style.bottom = 'auto';
}

function updateUntilNav(panel) {
  const { untilYear, untilMonth, minYear, minMonth, maxYear, maxMonth } = seriesPanelState;
  panel.querySelector('.sp-month-prev').disabled = untilYear === minYear && untilMonth === minMonth;
  panel.querySelector('.sp-month-next').disabled = untilYear === maxYear && untilMonth === maxMonth;
  panel.querySelector('.sp-month-label').textContent = spMonthLabel(untilYear, untilMonth);
}

function bindSeriesPanelEvents(panel, gearBtn) {
  // N −/+
  panel.querySelector('.sp-n-minus').addEventListener('click', async () => {
    if (seriesPanelState.n > seriesPanelState.minN) {
      seriesPanelState.n--;
      panel.querySelector('.sp-n-val').textContent = seriesPanelState.n;
      await updateSeriesPreview(panel);
    }
  });
  panel.querySelector('.sp-n-plus').addEventListener('click', async () => {
    if (seriesPanelState.n < 52) {
      seriesPanelState.n++;
      panel.querySelector('.sp-n-val').textContent = seriesPanelState.n;
      await updateSeriesPreview(panel);
    }
  });

  // Month nav ← →
  panel.querySelector('.sp-month-prev').addEventListener('click', async () => {
    const { untilYear, untilMonth, minYear, minMonth } = seriesPanelState;
    if (untilYear === minYear && untilMonth === minMonth) return;
    if (untilMonth === 0) { seriesPanelState.untilYear--; seriesPanelState.untilMonth = 11; }
    else seriesPanelState.untilMonth--;
    updateUntilNav(panel);
    await updateSeriesPreview(panel);
  });
  panel.querySelector('.sp-month-next').addEventListener('click', async () => {
    const { untilYear, untilMonth, maxYear, maxMonth } = seriesPanelState;
    if (untilYear === maxYear && untilMonth === maxMonth) return;
    if (untilMonth === 11) { seriesPanelState.untilYear++; seriesPanelState.untilMonth = 0; }
    else seriesPanelState.untilMonth++;
    updateUntilNav(panel);
    await updateSeriesPreview(panel);
  });

  updateUntilNav(panel);

  // Cancel / Close
  panel.querySelector('.sp-btn-cancel').addEventListener('click', () => {
    closeSeriesPanel();
    gearBtn.style.opacity = '';
  });

  // Insert / Undo
  panel.querySelector('.sp-btn-insert').addEventListener('click', async () => {
    if (seriesPanelState.insertedIds) {
      undoSeries(panel);
    } else {
      await insertSeries(panel);
    }
  });
}

async function computeSeriesCopies() {
  // Returns array of copies; each copy = array of { date, status }
  // preserving the full block shape of the source entry.
  const { id, startDate, sourceDates, n, type, untilYear, untilMonth } = seriesPanelState;
  const stepDays = n * 7;
  const untilEnd = new Date(untilYear, untilMonth + 1, 0); // last day of until-month

  // Day offsets of source block relative to block start
  const [sy, sm, sd] = startDate.split('-').map(Number);
  const blockStart = new Date(sy, sm - 1, sd);
  const offsets = sourceDates.map(dStr => {
    const [dy, dm, dd] = dStr.split('-').map(Number);
    return Math.round((new Date(dy, dm - 1, dd) - blockStart) / 86400000);
  });

  // Exclude own block dates to avoid self-collision
  const existingNonNote = new Set(
    absences.filter(a => a.type !== 'note' && a.id !== id).flatMap(a => a.dates)
  );
  const existingNotes = new Set(
    absences.filter(a => a.type === 'note' && a.id !== id).flatMap(a => a.dates)
  );

  const copies = [];
  let copyStart = new Date(blockStart);
  copyStart.setDate(copyStart.getDate() + stepDays);

  while (copyStart <= untilEnd) {
    const days = [];
    for (const offset of offsets) {
      const dayDate = new Date(copyStart);
      dayDate.setDate(dayDate.getDate() + offset);
      if (dayDate > untilEnd) continue;

      const y = dayDate.getFullYear();
      await ensureHolidays(y);
      const dStr = mkDate(dayDate.getFullYear(), dayDate.getMonth(), dayDate.getDate());
      const dow = dayDate.getDay();
      const isWE = dow === 0 || dow === 6;
      const holName = holidaysCache[y]?.[dStr];

      let status = 'ok';
      if (type === 'note') {
        if (existingNotes.has(dStr)) status = 'note';
      } else {
        if (isWE)          status = 'weekend';
        else if (holName)  status = 'holiday:' + holName;
        else if (existingNonNote.has(dStr)) status = 'occupied';
      }
      days.push({ date: dStr, status });
    }
    if (days.length > 0) copies.push(days);
    copyStart.setDate(copyStart.getDate() + stepDays);
  }

  return copies;
}

async function updateSeriesPreview(panel) {
  const preview = panel.querySelector('.sp-preview');
  const collDiv = panel.querySelector('.sp-collisions');
  const insertBtn = panel.querySelector('.sp-btn-insert');

  preview.innerHTML = '<div class="sp-preview-loading">Berechne…</div>';
  collDiv.style.display = 'none';
  insertBtn.disabled = true;

  const copies = await computeSeriesCopies();

  if (copies.length === 0) {
    preview.innerHTML = `<div class="sp-preview-empty">Keine Termine im gewählten Zeitraum.</div>`;
    return;
  }

  const allDays = copies.flat();
  const okDays = allDays.filter(d => d.status === 'ok');
  const skippedDays = allDays.filter(d => d.status !== 'ok');
  const okCopies = copies.filter(c => c.some(d => d.status === 'ok'));

  preview.innerHTML = `
    <div class="sp-preview-row">
      <span class="sp-preview-ok">✅ ${okCopies.length} Kopie${okCopies.length !== 1 ? 'n' : ''} (${okDays.length} Tage)</span>
      ${skippedDays.length > 0
        ? `<span class="sp-preview-warn">⏭️ ${skippedDays.length} übersprungen</span>`
        : '<span class="sp-preview-clean">Keine Kollisionen</span>'
      }
    </div>
  `;

  if (skippedDays.length > 0) {
    collDiv.style.display = 'block';
    collDiv.innerHTML = '<div class="sp-coll-title">Übersprungene Tage:</div>' +
      skippedDays.map(c => {
        const [y, m, d] = c.date.split('-').map(Number);
        const dateLabel = `${String(d).padStart(2,'0')}.${String(m).padStart(2,'0')}.${y}`;
        const dow = DAYS_SHORT[new Date(y, m-1, d).getDay()];
        let icon = '🚫', reason = '';
        if (c.status === 'weekend')               { icon = '📅'; reason = 'Wochenende'; }
        else if (c.status.startsWith('holiday:')) { icon = '🎉'; reason = c.status.slice(8); }
        else if (c.status === 'occupied')          { icon = '🚫'; reason = 'Bereits belegt'; }
        else if (c.status === 'note')              { icon = '📌'; reason = 'Bereits eine Notiz'; }
        return `<div class="sp-coll-item"><span class="sp-coll-icon">${icon}</span><span class="sp-coll-date">${dateLabel} ${dow}</span><span class="sp-coll-reason">${escapeHtml(reason)}</span></div>`;
      }).join('');
  }

  if (okDays.length > 0) {
    insertBtn.disabled = false;
    insertBtn.textContent = 'Einfügen';
    insertBtn.classList.remove('sp-btn-undo');
  }
}

async function insertSeries(panel) {
  const { id, type } = seriesPanelState;
  const copies = await computeSeriesCopies();

  // Carry over customLabel from source entry (note, other)
  const sourceAb = absences.find(a => a.id === id);
  const customLabel = sourceAb?.customLabel || '';

  const insertedIds = [];
  for (const copy of copies) {
    const okDates = copy.filter(d => d.status === 'ok').map(d => d.date);
    if (okDates.length === 0) continue;
    // Each copy = one independent entry (stamp logic, Option B)
    const newId = 'ab_' + Date.now() + '_' + Math.random().toString(36).slice(2, 7);
    const entry = { id: newId, type, dates: okDates.sort() };
    if (customLabel) entry.customLabel = customLabel;
    absences.push(entry);
    insertedIds.push(newId);
  }

  if (insertedIds.length === 0) {
    closeSeriesPanel();
    return;
  }

  seriesPanelState.insertedIds = insertedIds;
  saveData();
  renderMonth();
  updateStats();
  showSeriesDoneState(panel, insertedIds.length);
}

function showSeriesDoneState(panel, count) {
  const preview   = panel.querySelector('.sp-preview');
  const collDiv   = panel.querySelector('.sp-collisions');
  const insertBtn = panel.querySelector('.sp-btn-insert');
  const cancelBtn = panel.querySelector('.sp-btn-cancel');
  const rhythmRow = panel.querySelector('.sp-rhythm-row');
  const untilRow  = panel.querySelector('.sp-until-row');
  const blockInfo = panel.querySelector('.sp-block-info');

  if (rhythmRow) rhythmRow.style.display = 'none';
  if (untilRow)  untilRow.style.display  = 'none';
  if (blockInfo) blockInfo.style.display = 'none';
  collDiv.style.display = 'none';

  preview.innerHTML = `<div class="sp-preview-done">✅ ${count} Kopie${count !== 1 ? 'n' : ''} eingefügt</div>`;

  insertBtn.disabled = false;
  insertBtn.textContent = '⟲ Rückgängig';
  insertBtn.classList.add('sp-btn-undo');
  cancelBtn.textContent = 'Schließen';
}

function undoSeries(panel) {
  const { insertedIds } = seriesPanelState;
  if (!insertedIds) return;

  absences = absences.filter(a => !insertedIds.includes(a.id));
  seriesPanelState.insertedIds = null;

  saveData();
  renderMonth();
  updateStats();

  // Restore config state
  const rhythmRow = panel.querySelector('.sp-rhythm-row');
  const untilRow  = panel.querySelector('.sp-until-row');
  const blockInfo = panel.querySelector('.sp-block-info');
  const cancelBtn = panel.querySelector('.sp-btn-cancel');
  const insertBtn = panel.querySelector('.sp-btn-insert');

  if (rhythmRow) rhythmRow.style.display = '';
  if (untilRow)  untilRow.style.display  = '';
  if (blockInfo) blockInfo.style.display = '';
  cancelBtn.textContent = 'Abbrechen';
  insertBtn.classList.remove('sp-btn-undo');

  updateUntilNav(panel);
  updateSeriesPreview(panel);
}

// ─── Init sub-functions ──────────────────────────────────────────────────────

function initNavigation() {
  document.getElementById('prev-btn').onclick = () => {
    if (viewMode === 'year') {
      curYear--;
    } else {
      curMonth--;
      if (curMonth < 0) { curMonth = 11; curYear--; }
    }
    saveData();
    loadAndRender();
  };
  document.getElementById('next-btn').onclick = () => {
    if (viewMode === 'year') {
      curYear++;
    } else {
      curMonth++;
      if (curMonth > 11) { curMonth = 0; curYear++; }
    }
    saveData();
    loadAndRender();
  };
  document.getElementById('today-btn').onclick = () => {
    const n = new Date();
    curYear = n.getFullYear();
    curMonth = n.getMonth();
    saveData();
    loadAndRender();
  };
  document.getElementById('state-select').onchange = e => {
    currentState = e.target.value;
    holidaysCache = {};
    saveData();
    loadAndRender();
  };
  document.getElementById('view-toggle').onclick = () => {
    viewMode = viewMode === 'month' ? 'year' : 'month';
    saveData();
    loadAndRender();
  };
  document.getElementById('reset-btn').onclick = () => {
    if (confirm('⚠️ ALLE Einträge (Urlaub, Krank, Office-Marks) für diesen Monat wirklich LÖSCHEN?\n\nDas kann nicht rückgängig gemacht werden!')) {
      const prefix = `${curYear}-${pad(curMonth + 1)}-`;
      absences = absences.filter(ab => !ab.dates.some(d => d.startsWith(prefix)));
      saveData();
      loadAndRender();
    }
  };
}

function initDragDrop() {
  document.getElementById('chip-vacation').addEventListener('dragstart', e => e.dataTransfer.setData('type', 'vacation'));
  document.getElementById('chip-sick').addEventListener('dragstart', e => e.dataTransfer.setData('type', 'sick'));
  document.getElementById('chip-office').addEventListener('dragstart', e => e.dataTransfer.setData('type', 'office'));
  document.getElementById('chip-other').addEventListener('dragstart', e => e.dataTransfer.setData('type', 'other'));
  document.getElementById('chip-note')?.addEventListener('dragstart', e => e.dataTransfer.setData('type', 'note'));

  document.getElementById('cal-body').addEventListener('click', async e => {
    // Gear button → open series panel
    const gearBtn = e.target.closest('.gear-btn');
    if (gearBtn) {
      e.stopPropagation();
      const { id, date, type } = gearBtn.dataset;
      await openSeriesPanel(gearBtn, id, type, date);
      return;
    }

    const btn = e.target.closest('[data-id]');
    if (btn) {
      const id = btn.dataset.id;
      if (btn.classList.contains('btn-plus'))  await extendAbsence(id, 1);
      if (btn.classList.contains('btn-minus')) await extendAbsence(id, -1);
      if (btn.classList.contains('btn-del'))   deleteAbsence(id);
    }
    const noteBtn = e.target.closest('[data-note-id]');
    if (noteBtn) {
      const nid = noteBtn.dataset.noteId;
      if (noteBtn.classList.contains('btn-note-del'))   deleteNote(nid);
      if (noteBtn.classList.contains('btn-note-plus'))  await extendNote(nid, 1);
      if (noteBtn.classList.contains('btn-note-minus')) await extendNote(nid, -1);
    }
  });

  document.getElementById('cal-body').addEventListener('change', e => {
    if (e.target.classList.contains('ab-label-input')) {
      const id = e.target.dataset.id;
      const ab = absences.find(a => a.id === id);
      if (ab) { ab.customLabel = e.target.value.trim(); saveData(); renderMonth(); }
    }
    if (e.target.classList.contains('note-label-input')) {
      const id = e.target.dataset.noteId;
      const ab = absences.find(a => a.id === id);
      if (ab) { ab.customLabel = e.target.value.trim(); saveData(); renderMonth(); }
    }
  });
  document.getElementById('cal-body').addEventListener('keydown', e => {
    if ((e.target.classList.contains('ab-label-input') || e.target.classList.contains('note-label-input')) && e.key === 'Enter') {
      e.target.blur();
    }
  });
}

function initSettings() {
  const modal = document.getElementById('settings-modal');

  const savedModalSize = JSON.parse(localStorage.getItem('officeTrackerModalSize') || 'null');
  if (savedModalSize) {
    modal.querySelector('.settings-modal-inner').style.width = savedModalSize.w + 'px';
    modal.querySelector('.settings-modal-inner').style.height = savedModalSize.h + 'px';
  }

  const resizeHandle = document.getElementById('settings-resize-handle');
  const modalInner = modal.querySelector('.settings-modal-inner');
  let isResizing = false, startX, startY, startW, startH;

  resizeHandle.addEventListener('mousedown', e => {
    isResizing = true;
    startX = e.clientX; startY = e.clientY;
    startW = modalInner.offsetWidth; startH = modalInner.offsetHeight;
    document.body.style.userSelect = 'none';
    e.preventDefault();
  });

  document.addEventListener('mousemove', e => {
    if (!isResizing) return;
    const newW = Math.max(600, Math.min(window.innerWidth * 0.95, startW + e.clientX - startX));
    const newH = Math.max(400, Math.min(window.innerHeight * 0.92, startH + e.clientY - startY));
    modalInner.style.width = newW + 'px';
    modalInner.style.height = newH + 'px';
  });

  document.addEventListener('mouseup', () => {
    if (!isResizing) return;
    isResizing = false;
    document.body.style.userSelect = '';
    localStorage.setItem('officeTrackerModalSize', JSON.stringify({
      w: modalInner.offsetWidth,
      h: modalInner.offsetHeight
    }));
    modal._suppressClose = true;
    setTimeout(() => { modal._suppressClose = false; }, 50);
  });

  const navItems = modal.querySelectorAll('.settings-nav-item');
  const panels = modal.querySelectorAll('.settings-panel');

  function switchPanel(panelId) {
    navItems.forEach(n => n.classList.remove('active'));
    panels.forEach(p => p.classList.remove('active'));
    modal.querySelector(`[data-panel="${panelId}"]`).classList.add('active');
    modal.querySelector(`#panel-${panelId}`).classList.add('active');
  }

  navItems.forEach(item => {
    item.addEventListener('click', () => switchPanel(item.dataset.panel));
  });

  document.getElementById('settings-btn').onclick = () => {
    modal.style.display = 'flex';
    switchPanel('work');
    requestAnimationFrame(() => modal.classList.add('visible'));
    document.getElementById('hours-per-week').value = settings.hoursPerWeek;
    settings.workDays.forEach(d => {
      const el = document.getElementById(`workday-${d}`);
      if (el) el.checked = true;
    });
    COLOR_KEYS.forEach(c => {
      const el = document.getElementById(`color-${c}`);
      if (el) el.value = settings.colors[c];
    });
    renderCustomHolidays();
  };

  const closeModal = () => {
    modal.classList.remove('visible');
    setTimeout(() => { modal.style.display = 'none'; }, 260);
  };
  document.getElementById('close-settings').onclick = closeModal;
  modal.addEventListener('click', e => { if (e.target === modal && !modal._suppressClose) closeModal(); });

  document.getElementById('hours-per-week').onchange = e => {
    const val = parseFloat(e.target.value) || 40;
    settings.hoursPerWeek = Math.max(1, Math.min(60, val));
    e.target.value = settings.hoursPerWeek;
    saveData();
    updateStats();
  };

  [0, 1, 2, 3, 4, 5, 6].forEach(d => {
    const el = document.getElementById(`workday-${d}`);
    if (el) {
      el.onchange = () => {
        settings.workDays = [0,1,2,3,4,5,6].filter(x => document.getElementById(`workday-${x}`)?.checked);
        if (settings.workDays.length === 0) settings.workDays = [1, 2, 3, 4, 5];
        saveData();
        holidaysCache = {};
        loadAndRender();
      };
    }
  });

  COLOR_KEYS.forEach(c => {
    const el = document.getElementById(`color-${c}`);
    if (el) {
      el.oninput = e => { settings.colors[c] = e.target.value; applyCustomColors(); saveData(); };
    }
  });

  document.getElementById('reset-colors').onclick = () => {
    settings.colors = JSON.parse(JSON.stringify(DEFAULT_COLORS));
    applyCustomColors();
    saveData();
    COLOR_KEYS.forEach(c => {
      const el = document.getElementById(`color-${c}`);
      if (el) el.value = settings.colors[c];
    });
  };
}

function initCustomHolidays() {
  document.getElementById('custom-holidays-list').addEventListener('click', e => {
    const btn = e.target.closest('.remove-holiday-btn');
    if (btn) {
      delete settings.customHolidays[btn.dataset.date];
      saveData();
      holidaysCache = {};
      renderCustomHolidays();
      loadAndRender();
    }
  });

  document.getElementById('add-custom-holiday').onclick = () => {
    const dateInput = document.getElementById('custom-holiday-date');
    const nameInput = document.getElementById('custom-holiday-name');
    const date = dateInput.value;
    const name = nameInput.value.trim();
    if (date && name) {
      settings.customHolidays[date] = name;
      saveData();
      holidaysCache = {};
      dateInput.value = '';
      nameInput.value = '';
      renderCustomHolidays();
      loadAndRender();
    }
  };
}

function initDataManagement() {
  document.getElementById('export-data').onclick = () => {
    const data = { absences, settings, view: { year: curYear, month: curMonth, state: currentState, mode: viewMode } };
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `office-tracker-backup-${new Date().toISOString().split('T')[0]}.json`;
    a.click();
    URL.revokeObjectURL(url);
  };

  document.getElementById('import-data').onclick = () => document.getElementById('import-file').click();

  document.getElementById('import-file').onchange = e => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const data = JSON.parse(ev.target.result);
        validateImport(data);
        if (data.absences) absences = data.absences;
        if (data.settings) { settings = { ...settings, ...data.settings }; applyCustomColors(); }
        if (data.view) {
          curYear = typeof data.view.year === 'number' && data.view.year >= 2000 && data.view.year <= 2100 ? data.view.year : new Date().getFullYear();
          curMonth = typeof data.view.month === 'number' && data.view.month >= 0 && data.view.month <= 11 ? data.view.month : new Date().getMonth();
          currentState = data.view.state || 'DE-BW';
          viewMode = data.view.mode || 'month';
        }
        holidaysCache = {};
        saveData();
        loadAndRender();
        document.getElementById('settings-modal').classList.remove('visible');
        setTimeout(() => { document.getElementById('settings-modal').style.display = 'none'; }, 260);
      } catch (err) {
        alert('Import fehlgeschlagen: ' + err.message);
      }
    };
    reader.readAsText(file);
    e.target.value = '';
  };

  document.getElementById('clear-all-data').onclick = () => {
    if (confirm('⚠️ WIRKLICH ALLE DATEN LÖSCHEN? Dies kann nicht rückgängig gemacht werden!')) {
      localStorage.removeItem('officeTracker');
      location.reload();
    }
  };
}

// ─── Entry point ─────────────────────────────────────────────────────────────

function init() {
  const now = new Date();
  curYear = now.getFullYear();
  curMonth = now.getMonth();

  loadData();

  const setStickyHeight = () => {
    const sticky = document.querySelector('.sticky-top');
    if (sticky) {
      document.documentElement.style.setProperty('--sticky-height', sticky.offsetHeight + 'px');
    }
  };
  setStickyHeight();
  window.addEventListener('resize', setStickyHeight);

  document.getElementById('state-select').value = currentState;

  initNavigation();
  initDragDrop();
  initSettings();
  initCustomHolidays();
  initDataManagement();

  loadAndRender();
}

init();
