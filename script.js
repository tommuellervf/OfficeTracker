const MONTHS_DE = ['Januar','Februar','März','April','Mai','Juni','Juli','August','September','Oktober','November','Dezember'];
const DAYS_SHORT = ['So','Mo','Di','Mi','Do','Fr','Sa'];

let curYear, curMonth;
let currentState = 'DE-BW';
let holidaysCache = {};
let absences = [];
let viewMode = 'month';
let settings = {
  hoursPerWeek: 40,
  workDays: [1,2,3,4,5],
  customHolidays: {},
  colors: {
    accent: '#2563eb',
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
    bg: '#eef4fd'
  }
};

const DEFAULT_COLORS = JSON.parse(JSON.stringify(settings.colors));

const pad = n => String(n).padStart(2,'0');
const mkDate = (y,m,d) => `${y}-${pad(m+1)}-${pad(d)}`;

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
  root.style.setProperty('--accent2', settings.colors.accent);
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
    const url = `https://openholidaysapi.org/PublicHolidays?countryIsoCode=DE&subdivisionCode=${subdivisionCode}&languageIsoCode=DE&validFrom=${year}-01-01&validTo=${year}-12-31`;
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
  renderHolidayList();
  updateStats();

  document.getElementById('view-toggle').textContent = viewMode === 'month' ? 'Jahr' : 'Monat';
}

function isHoliday(dateStr) {
  const year = dateStr.split('-')[0];
  return !!holidaysCache[year]?.[dateStr];
}

function renderMonth() {
  const body = document.getElementById('cal-body');
  const now = new Date();
  const todayStr = mkDate(now.getFullYear(), now.getMonth(), now.getDate());
  const isCurrentMonth = (curYear === now.getFullYear() && curMonth === now.getMonth());
  const dim = new Date(curYear, curMonth+1, 0).getDate();
  let lastKW = null;
  body.innerHTML = '';

  for(let d=1; d<=dim; d++){
    const date = new Date(curYear, curMonth, d);
    const dow = date.getDay();
    const isWE = dow===0 || dow===6;
    const isWork = isWorkday(dow);
    const dateStr = mkDate(curYear, curMonth, d);
    const isToday = dateStr === todayStr && isCurrentMonth;
    const hol = !!holidaysCache[curYear]?.[dateStr];
    const absInfo = getAbsInfo(dateStr);
    const isOffice = absInfo && absInfo.ab.type === 'office';
    const isOther  = absInfo && absInfo.ab.type === 'other';
    const isVacation = absInfo && absInfo.ab.type === 'vacation';
    const isSick = absInfo && absInfo.ab.type === 'sick';
    const kw = getWeekNum(date);
    const showKW = isWork && kw !== lastKW;
    if(isWork) lastKW = kw;

    const row = document.createElement('div');
    row.className = 'day-row'
      + (isWE ? ' weekend' : '')
      + (isToday ? ' is-today' : '')
      + (isOffice ? ' office-day' : '')
      + (isVacation ? ' vacation-day' : '')
      + (isSick ? ' sick-day' : '')
      + (isOther ? ' other-day' : '')
      + (isWork && !hol && !absInfo ? ' droppable' : '');
    row.dataset.date = dateStr;

    const kwCell = document.createElement('div');
    kwCell.className = 'kw-cell';
    kwCell.textContent = showKW ? `${kw}` : '';

    const dayCell = document.createElement('div');
    dayCell.className = 'day-cell';
    dayCell.innerHTML = `<span class="day-num-big">${d}</span><span class="day-name-small">${DAYS_SHORT[dow]}</span>${isToday ? '<div class="today-badge"></div>' : ''}`;

    const absCell = document.createElement('div');
    absCell.className = 'absence-cell';

    if(hol) {
      const bl = document.createElement('div');
      bl.className = 'ab-block holiday pos-single';
      bl.innerHTML = `<span class="ab-label">🎉 ${holidaysCache[curYear][dateStr]}</span>`;
      absCell.appendChild(bl);
    } else if(absInfo) {
      const {ab, idx} = absInfo;
      const total = ab.dates.length;
      const bl = document.createElement('div');
      bl.className = `ab-block ${ab.type}`;

      const emoji = ab.type==='vacation' ? '🌴' : ab.type==='sick' ? '😷' : ab.type==='office' ? '🏢' : '📋';
      const label = ab.type==='vacation' ? 'Urlaub' : ab.type==='sick' ? 'Krank' : ab.type==='office' ? 'Im Office' : (ab.customLabel || 'Sonstige Abw.');
      const labelHtml = ab.type === 'other'
        ? idx === 0
          ? `<input class="ab-label-input" data-id="${ab.id}" value="${(ab.customLabel||'').replace(/"/g,'&quot;')}" placeholder="Sonstige Abw." title="Bezeichnung ändern">`
          : `<span class="ab-label" style="opacity:0.6;font-size:10px;">${ab.customLabel || 'Sonstige'}</span>`
        : `<span class="ab-label">${emoji} ${label}</span>`;
      bl.innerHTML = `
        ${labelHtml}
        ${idx === 0 ? `
        <div class="ab-controls">
          <button class="ctrl-btn btn-minus" data-id="${ab.id}">−</button>
          <span class="ab-day-count">${total}d</span>
          <button class="ctrl-btn btn-plus" data-id="${ab.id}">+</button>
          <button class="ctrl-del btn-del" data-id="${ab.id}" style="background:none;border:none;cursor:pointer;display:flex;align-items:center;justify-content:center;padding:0;"><svg width="10" height="10" viewBox="0 0 14 14" fill="none"><path d="M1 1L13 13M1 13L13 1" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg></button>
        </div>` : ''}`;
      absCell.appendChild(bl);
    }

    row.appendChild(kwCell);
    row.appendChild(dayCell);
    row.appendChild(absCell);
    body.appendChild(row);

    if(isWork && !hol && !absInfo) {
      row.addEventListener('dragover', e => { e.preventDefault(); row.classList.add('drag-over'); });
      row.addEventListener('dragleave', () => row.classList.remove('drag-over'));
      row.addEventListener('drop', e => {
        e.preventDefault();
        row.classList.remove('drag-over');
        const type = e.dataTransfer.getData('type');
        if(type) addEntry(type, dateStr);
      });
    }
  }
}

async function renderYearOverview() {
  await ensureHolidays(curYear);

  const body = document.getElementById('cal-body');
  body.innerHTML = '';
  const grid = document.createElement('div');
  grid.className = 'year-grid';

  const now = new Date();
  const todayMonth = now.getMonth();
  const todayYear = now.getFullYear();
  const baseDays = Math.round(settings.hoursPerWeek / 5);

  for (let m = 0; m < 12; m++) {
    const prefix = `${curYear}-${pad(m+1)}-`;
    const isCurrent = curYear === todayYear && m === todayMonth;
    const isPast    = curYear < todayYear || (curYear === todayYear && m < todayMonth);
    const isFuture  = !isCurrent && !isPast;

    const dim = new Date(curYear, m+1, 0).getDate();
    let workdays = 0, holidayCount = 0;
    for (let d = 1; d <= dim; d++) {
      const dow = new Date(curYear, m, d).getDay();
      if (isWorkday(dow)) {
        workdays++;
        if (holidaysCache[curYear]?.[mkDate(curYear, m, d)]) holidayCount++;
      }
    }

    const vacDays  = getAbsencesInMonth(m, 'vacation');
    const sickDays = getAbsencesInMonth(m, 'sick');
    const otherDays = getAbsencesInMonth(m, 'other');
    const attended = getAttendedInMonth(m);

    const allAbs = vacDays + sickDays + otherDays;
    const effectiveWorkdays = workdays - holidayCount;
    const raw = effectiveWorkdays === 0 ? baseDays : baseDays - (allAbs * (baseDays / effectiveWorkdays));
    const pflicht = Math.round(Math.max(0, raw));

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

    let dotsHtml = '';
    for (let d = 1; d <= dim; d++) {
      const dow = new Date(curYear, m, d).getDay();
      const dStr = mkDate(curYear, m, d);
      const isWE = dow === 0 || dow === 6;
      const absI = getAbsInfo(dStr);
      const isHol = !!holidaysCache[curYear]?.[dStr];
      let dotClass = 'workday';
      if (isWE) dotClass = 'weekend';
      else if (isHol) dotClass = 'holiday';
      else if (absI) dotClass = absI.ab.type;
      dotsHtml += `<div class="year-dot ${dotClass}" title="${dStr}"></div>`;
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

function getAbsencesInMonth(m, type = null) {
  const prefix = `${curYear}-${pad(m+1)}-`;
  return absences
    .filter(ab => type === null || ab.type === type)
    .reduce((sum, ab) => sum + ab.dates.filter(d => d.startsWith(prefix)).length, 0);
}

function getAttendedInMonth(m) {
  return getAbsencesInMonth(m, 'office');
}

async function addEntry(type, startDate) {
  const dates = await computeDates(startDate, 1);
  if (dates.length) absences.push({ id: 'ab_'+Date.now(), type, dates });
  saveData();
  renderMonth();
  updateStats();
}

function getAbsInfo(dateStr) {
  for(const ab of absences){
    const idx = ab.dates.indexOf(dateStr);
    if(idx !== -1) return {ab, idx};
  }
  return null;
}

function getWeekNum(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay()||7));
  const y = new Date(Date.UTC(d.getUTCFullYear(),0,1));
  return Math.ceil((((d-y)/86400000)+1)/7);
}

async function computeDates(startDate, numDays) {
  const [sy,sm,sd] = startDate.split('-').map(Number);
  const existing = new Set(absences.flatMap(a => a.dates));
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

const MAX_FUTURE_DAYS = 180;

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

    const existing = new Set(absences.flatMap(a => a.dates));

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
  absences = absences.filter(a => a.id !== id);
  saveData();
  renderMonth();
  updateStats();
}

function updateStats() {
  const dim = new Date(curYear, curMonth+1, 0).getDate();
  let workdays=0, holidayCount=0;
  for(let d=1; d<=dim; d++){
    const dow = new Date(curYear,curMonth,d).getDay();
    if(isWorkday(dow)){ 
      workdays++; 
      if(isHoliday(mkDate(curYear,curMonth,d))) holidayCount++; 
    }
  }
  const prefix = `${curYear}-${pad(curMonth+1)}-`;
  const stats = absences.reduce((acc, ab) => {
    const count = ab.dates.filter(d => d.startsWith(prefix)).length;
    if (count === 0) return acc;
    if (ab.type === 'office') {
      acc.attended += count;
    } else if (ab.type === 'vacation' || ab.type === 'sick' || ab.type === 'other') {
      acc.absCount += count;
    }
    return acc;
  }, { absCount: 0, attended: 0 });

  const allAbs = stats.absCount;
  const netWorkdays = workdays - holidayCount - stats.absCount;
  const baseDays = Math.round(settings.hoursPerWeek / 5);
  const effectiveWorkdays = workdays - holidayCount;
  const raw = effectiveWorkdays === 0 ? baseDays : baseDays - (allAbs * (baseDays / effectiveWorkdays));
  const officePflicht = Math.round(Math.max(0, raw));

  document.getElementById('stat-workdays').textContent = workdays;
  document.getElementById('stat-holidays').textContent = holidayCount;
  document.getElementById('stat-absences').textContent = stats.absCount;
  document.getElementById('stat-net').textContent = Math.max(0, netWorkdays);
  document.getElementById('stat-basedays').textContent = baseDays;
  document.getElementById('stat-basedays-sub').textContent = `${settings.hoursPerWeek}h ÷ 5 Tage/Woche`;

  // ── Office-Card: Status-Klasse setzen ──────────────────────────────────────
  const officeCard = document.getElementById('stat-office').closest('.stat-card');
  if (officeCard) {
    officeCard.classList.remove('status-ok', 'status-warn');
    if (officePflicht > 0) {
      officeCard.classList.add(stats.attended >= officePflicht ? 'status-ok' : 'status-warn');
    }
  }
  document.getElementById('stat-office').innerHTML =
    `<span>${stats.attended}</span> / ${officePflicht}`;
}

function renderHolidayList() {
  const list = document.getElementById('holiday-list');
  const items = [];
  for(let d=1; d<= new Date(curYear, curMonth+1, 0).getDate(); d++){
    const dStr = mkDate(curYear, curMonth, d);
    if(holidaysCache[curYear]?.[dStr]) items.push({date:dStr, name:holidaysCache[curYear][dStr]});
  }
  if(!items.length){
    list.innerHTML = `<div class="holiday-empty">Keine Feiertage</div>`;
    return;
  }
  list.innerHTML = items.map(h => {
    const p = h.date.split('-');
    return `<div class="holiday-chip"><span>🎉 ${h.name}</span><span class="holiday-chip-date">${p[2]}.${p[1]}.</span></div>`;
  }).join('');
}

// ─── Helpers ────────────────────────────────────────────────────────────────

function escapeHtml(str) {
  return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

const DATE_RE = /^\d{4}-\d{2}-\d{2}$/;
const VALID_TYPES = ['vacation', 'sick', 'office', 'other'];
const COLOR_KEYS = ['accent', 'vacation', 'vacation-bg', 'vacation-border', 'sick', 'sick-bg', 'sick-border', 'holiday', 'holiday-bg', 'holiday-border', 'office', 'office-bg', 'office-border', 'other', 'other-bg', 'other-border', 'bg'];

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

  document.getElementById('cal-body').addEventListener('click', async e => {
    const btn = e.target.closest('[data-id]');
    if (btn) {
      const id = btn.dataset.id;
      if (btn.classList.contains('btn-plus'))  await extendAbsence(id, 1);
      if (btn.classList.contains('btn-minus')) await extendAbsence(id, -1);
      if (btn.classList.contains('btn-del'))   deleteAbsence(id);
    }
  });

  document.getElementById('cal-body').addEventListener('change', e => {
    if (e.target.classList.contains('ab-label-input')) {
      const id = e.target.dataset.id;
      const ab = absences.find(a => a.id === id);
      if (ab) { ab.customLabel = e.target.value.trim(); saveData(); renderMonth(); }
    }
  });
  document.getElementById('cal-body').addEventListener('keydown', e => {
    if (e.target.classList.contains('ab-label-input') && e.key === 'Enter') {
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
        settings.workDays = [0, 1, 2, 3, 4, 5, 6].filter(x => document.getElementById(`workday-${x}`).checked);
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
