/* Cashew — разбор банковской выписки и аналитика расходов. Всё считается в браузере. */

'use strict';

const STORE_KEY = 'cashew.tx.v1';
const THEME_KEY = 'cashew.theme';

/* ---------- категории ---------- */

const CATEGORIES = [
  { name: 'Транспорт',          color: '#60a5fa', rx: /TELCELL TRANSPORT|YANDEX\.? ?GO|GG TAXI|UTAXI|METRO|BUS/i },
  { name: 'Продукты',           color: '#34d399', rx: /ORANGE MARKET|APRICOT MARKET|SAS |SAS-|YEREVAN CITY|SUPERMARKET|PRIME FARM|HAYR U VORDI|MARKET|GROCER|CARREFOUR|EVRIKA/i },
  { name: 'Кафе и рестораны',   color: '#fb923c', rx: /COFFEE|KFC|SANDWICH|FOODICTED|YANDEX EATS|MOKA POT|CAFE|PIZZA|BURGER|RESTAURANT|BAR |DODO|TASHIR/i },
  { name: 'Подписки и игры',    color: '#a78bfa', rx: /STEAM|DISCORD|SPOTIFY|NETFLIX|APPLE\.COM|GOOGLE ?\*|YOUTUBE|PATREON|OPENAI|CLAUDE|ANTHROPIC/i },
  { name: 'Связь и интернет',   color: '#22d3ee', rx: /UCOM|VIVA|TEAM ?TELECOM|BEELINE|ROSTELECOM|INTERNET/i },
  { name: 'Покупки',            color: '#f472b6', rx: /WILDBERRIES|OZON|ALIEXPRESS|ZARA|LC ?WAIKIKI|H&M|WL ?\*|SHOP|STORE|GLIDE ADVENTURES|CJ /i },
  { name: 'Здоровье и красота', color: '#fb7185', rx: /PHARM|GEDEON RICHTER|APTEK|DENTAL|CLINIC|MEDIC|NATALI|BEAUTY|SALON/i },
  { name: 'Переводы и счета',   color: '#64748b', rx: /փոխանցում|համալրում|TRANSFER|CASH|ATM|Կանխիկ/i },
  { name: 'Комиссии и налоги',  color: '#eab308', rx: /Հարկ|Tax|միջնորդավճար|COMMISSION|FEE/i },
];
const OTHER = { name: 'Другое', color: '#8b95b8' };

function categorize(desc) {
  for (const c of CATEGORIES) if (c.rx.test(desc)) return c.name;
  return OTHER.name;
}
function catColor(name) {
  const c = CATEGORIES.find(x => x.name === name);
  return c ? c.color : OTHER.color;
}

/* ---------- разбор файла ---------- */

const DATE_RX = /^(\d{2})\.(\d{2})\.(\d{4})$/;
const PLUS_RX = /^\+\s*[\d\s  .,]+$/;
const MINUS_RX = /^[-–—−]\s*[\d\s  .,]+$/;

function toNumber(raw) {
  let s = String(raw).replace(/[\s  +\-–—−]/g, '');
  if (/,\d{1,2}$/.test(s)) s = s.replace(/\./g, '').replace(',', '.');
  else s = s.replace(/,/g, '');
  const n = parseFloat(s);
  return isFinite(n) ? n : 0;
}

// Из строки-описания достаём читаемое имя места.
function merchantOf(desc) {
  let s = String(desc)
    .replace(/^Ք:\s*/, '')
    .split('Ամերիաբանկ')[0]
    .split('ՀԾ-')[0]
    .split('ՀԱՇՎԱՐԿԱՅԻՆ')[0]
    .trim();
  if (/^%\s*կապիտ/i.test(s)) return 'Проценты по счёту';
  if (/Հարկ/i.test(s)) return 'Налог';
  if (/փոխանցում/i.test(s)) return 'Личный перевод';
  if (/համալրում/i.test(s)) return 'Пополнение другого счёта';
  s = s.split(/\s*[:/]\s*/)[0];                        // "Ucom Payment: 0950..." → "Ucom Payment"
  s = s.split(/\s+(?:YEREVAN|Yerevan|EREVAN)\b/)[0];   // отрезаем город и код терминала
  s = s.replace(/(\s+\d[\d-]*)+\s*$/, '').replace(/\s+(?:AM|LLC|Payment)$/i, '').replace(/\s+/g, ' ').trim();
  s = s.replace(/\b(\S+)(\s+\1\b)+/gi, '$1');          // "EON EON" → "EON"
  return s || String(desc).slice(0, 28);
}

function parseWorkbook(data, fileName) {
  const wb = XLSX.read(data, { type: 'array' });
  const rows = [];
  for (const sheetName of wb.SheetNames) {
    const grid = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, raw: false, defval: '' });
    for (const row of grid) {
      const cells = row.map(c => (c == null ? '' : String(c).trim()));
      const dateCell = cells.find(c => DATE_RX.test(c));
      if (!dateCell) continue;

      const plus = cells.find(c => PLUS_RX.test(c));
      const minus = cells.find(c => MINUS_RX.test(c));
      if (!plus && !minus) continue;

      // описание — самая «текстовая» ячейка строки
      let desc = '';
      for (const c of cells) {
        if (DATE_RX.test(c) || PLUS_RX.test(c) || MINUS_RX.test(c)) continue;
        if (c.length <= 3 || /^[\d\s.,+-]+$/.test(c)) continue;
        if (c.length > desc.length) desc = c;
      }
      if (!desc) continue;

      const [, dd, mm, yyyy] = dateCell.match(DATE_RX);
      rows.push({
        date: `${yyyy}-${mm}-${dd}`,
        amount: toNumber(plus || minus),
        type: plus ? 'in' : 'out',
        desc,
        merchant: merchantOf(desc),
        cat: categorize(desc),
        file: fileName,
      });
    }
  }

  // Одинаковые операции в один день бывают настоящими (два раза по 150 за транспорт),
  // поэтому в ключ идёт ещё и номер повтора внутри файла — так пересекающиеся
  // выписки схлопываются, а честные дубли остаются.
  const seen = new Map();
  for (const r of rows) {
    const base = `${r.date}|${r.type}|${r.amount}|${r.desc}`;
    const n = (seen.get(base) || 0) + 1;
    seen.set(base, n);
    r.uid = `${base}|${n}`;
  }
  return rows;
}

/* ---------- состояние ---------- */

let TX = [];       // все операции
let VIEW = [];     // операции выбранного периода
let shown = 50;
let selectedDay = null;  // день, раскрытый под графиком

const $ = id => document.getElementById(id);
const fmt = n => new Intl.NumberFormat('ru-RU', { maximumFractionDigits: 0 }).format(Math.round(n)) + ' ֏';
// «1 операция / 2 операции / 5 операций»
const plural = (n, one, few, many) => {
  const m10 = n % 10, m100 = n % 100;
  if (m10 === 1 && m100 !== 11) return `${n} ${one}`;
  if (m10 >= 2 && m10 <= 4 && (m100 < 12 || m100 > 14)) return `${n} ${few}`;
  return `${n} ${many}`;
};
const fmtDate = iso => iso.slice(8, 10) + '.' + iso.slice(5, 7) + '.' + iso.slice(0, 4);
const monthName = key => {
  const names = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];
  const [y, m] = key.split('-');
  return names[+m - 1] + ' ' + y;
};

/* ---------- загрузка файлов ---------- */

function readFiles(fileList) {
  const files = [...fileList];
  if (!files.length) return;
  $('uploadError').classList.add('hidden');

  Promise.all(files.map(f => f.arrayBuffer().then(buf => {
    try { return parseWorkbook(new Uint8Array(buf), f.name); }
    catch (e) { console.error(e); return []; }
  }))).then(results => {
    const rows = results.flat();
    if (!rows.length) {
      showUploadError('Не удалось найти операции в файле. Загрузите выписку банка в формате .xlsx — с колонками «Дата», «Мутք/Ելք» и описанием операции.');
      return;
    }

    // новые выписки добавляются к уже загруженным; повторы отбрасываем
    const known = new Set(TX.map(t => t.uid));
    const fresh = rows.filter(t => !known.has(t.uid));
    fresh.forEach(t => known.add(t.uid));

    if (!fresh.length) {
      showUploadError('Эти операции уже загружены — выписка за этот период уже есть.');
      return;
    }

    TX = TX.concat(fresh).sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : 0));
    save();
    showDashboard();
    flash(rows.length - fresh.length
      ? `Добавлено ${plural(fresh.length, 'операция', 'операции', 'операций')}, ${rows.length - fresh.length} уже были загружены`
      : `Добавлено ${plural(fresh.length, 'операция', 'операции', 'операций')}`);
  });
}

function showUploadError(text) {
  const el = $('uploadError');
  el.textContent = text;
  el.classList.remove('hidden');
  flash(text);
}

// короткое сообщение над списком выписок
let flashTimer = 0;
function flash(text) {
  const el = $('fileNote');
  el.textContent = text;
  el.classList.remove('hidden');
  clearTimeout(flashTimer);
  flashTimer = setTimeout(() => el.classList.add('hidden'), 6000);
}

function save() {
  try { localStorage.setItem(STORE_KEY, JSON.stringify(TX)); }
  catch (e) { console.error('Не удалось сохранить в localStorage', e); }
}

/* ---------- список загруженных выписок ---------- */

function renderFiles() {
  const names = [...new Set(TX.map(t => t.file))];
  $('fileList').innerHTML = names.map(name => {
    const rows = TX.filter(t => t.file === name);
    const dates = rows.map(t => t.date).sort();
    const months = [...new Set(dates.map(d => d.slice(0, 7)))];
    const label = months.length === 1 ? monthName(months[0])
      : `${fmtDate(dates[0])} — ${fmtDate(dates[dates.length - 1])}`;
    return `
      <li class="file-chip">
        <span class="file-chip-main">
          <span class="file-chip-name">${esc(label)}</span>
          <span class="file-chip-meta">${plural(rows.length, 'операция', 'операции', 'операций')} · ${esc(name)}</span>
        </span>
        <button type="button" class="file-chip-x" data-file="${esc(name)}"
                title="Убрать эту выписку" aria-label="Убрать выписку ${esc(name)}">✕</button>
      </li>`;
  }).join('');
}

function removeFile(name) {
  TX = TX.filter(t => t.file !== name);
  save();
  if (!TX.length) { resetAll(); return; }
  showDashboard();
  flash('Выписка убрана');
}

/* ---------- отрисовка ---------- */

function showDashboard() {
  $('uploadView').classList.add('hidden');
  $('dashView').classList.remove('hidden');
  $('resetBtn').classList.remove('hidden');

  const months = [...new Set(TX.map(t => t.date.slice(0, 7)))].sort().reverse();
  const sel = $('monthSelect');
  const wanted = sel.value;                    // после добавления файла период не сбрасываем
  sel.innerHTML = '';
  if (months.length > 1) sel.append(new Option('Все периоды', 'all'));
  months.forEach(m => sel.append(new Option(monthName(m), m)));
  sel.value = [...sel.options].some(o => o.value === wanted) ? wanted
    : (months.length > 1 ? 'all' : months[0]);

  renderFiles();

  const cats = [...new Set(TX.filter(t => t.type === 'out').map(t => t.cat))].sort();
  const cf = $('catFilter');
  cf.innerHTML = '';
  cf.append(new Option('Все категории', 'all'));
  cats.forEach(c => cf.append(new Option(c, c)));

  render();
}

function render() {
  const m = $('monthSelect').value;
  VIEW = m === 'all' ? TX.slice() : TX.filter(t => t.date.slice(0, 7) === m);
  shown = 50;
  selectedDay = null;

  const outs = VIEW.filter(t => t.type === 'out');
  const ins = VIEW.filter(t => t.type === 'in');
  const totalOut = outs.reduce((s, t) => s + t.amount, 0);
  const totalIn = ins.reduce((s, t) => s + t.amount, 0);

  const dates = VIEW.map(t => t.date).sort();
  $('periodTitle').textContent = m === 'all' ? 'Все периоды' : monthName(m);
  $('periodSub').textContent = dates.length
    ? `${fmtDate(dates[0])} — ${fmtDate(dates[dates.length - 1])} · ${plural(VIEW.length, 'операция', 'операции', 'операций')}`
    : 'Нет операций';

  $('sumIncome').textContent = fmt(totalIn);
  $('cntIncome').textContent = plural(ins.length, 'поступление', 'поступления', 'поступлений');
  $('sumExpense').textContent = fmt(totalOut);
  $('cntExpense').textContent = plural(outs.length, 'трата', 'траты', 'трат');

  const net = totalIn - totalOut;
  const netEl = $('sumNet');
  netEl.textContent = (net >= 0 ? '+' : '−') + fmt(Math.abs(net)).replace('-', '');
  netEl.className = 'card-value ' + (net >= 0 ? 'income' : 'expense');
  $('netNote').textContent = net >= 0 ? 'отложено за период' : 'потрачено больше, чем пришло';

  const days = new Set(outs.map(t => t.date)).size || 1;
  $('avgDay').textContent = fmt(totalOut / days);
  $('avgNote').textContent = `за ${plural(days, 'день', 'дня', 'дней')} с тратами`;

  renderCategories(outs, totalOut);
  renderMerchants(outs);
  renderDays(outs);
  renderDayDetail();
  renderTable();
}

function groupSum(items, keyFn) {
  const map = new Map();
  for (const t of items) map.set(keyFn(t), (map.get(keyFn(t)) || 0) + t.amount);
  return [...map.entries()].map(([key, sum]) => ({ key, sum })).sort((a, b) => b.sum - a.sum);
}

function renderCategories(outs, total) {
  const groups = groupSum(outs, t => t.cat);
  const top = groups[0];

  // в центре кольца — доля самой крупной категории
  $('donutPct').textContent = top ? Math.round(top.sum / total * 100) + '%' : '—';
  $('donutCap').textContent = top ? top.key : 'нет расходов';

  const donut = $('donut');
  donut.innerHTML = '<circle cx="21" cy="21" r="15.915" fill="none" stroke="var(--track)" stroke-width="5"></circle>';
  let offset = 0;
  groups.forEach(g => {
    const pct = total ? (g.sum / total) * 100 : 0;
    const c = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
    c.setAttribute('cx', '21'); c.setAttribute('cy', '21'); c.setAttribute('r', '15.915');
    c.setAttribute('fill', 'none');
    c.setAttribute('stroke', catColor(g.key));
    c.setAttribute('stroke-width', '5');
    c.setAttribute('stroke-linecap', 'butt');
    c.setAttribute('stroke-dasharray', `0 100`);
    c.setAttribute('stroke-dashoffset', String(-offset));
    donut.appendChild(c);
    requestAnimationFrame(() => c.setAttribute('stroke-dasharray', `${pct} ${100 - pct}`));
    offset += pct;
  });

  $('catList').innerHTML = groups.map(g => `
    <li>
      <span class="dot" style="background:${catColor(g.key)}"></span>
      <span>${esc(g.key)} <span class="cat-pct">${Math.round(g.sum / total * 100)}%</span></span>
      <span class="cat-sum">${fmt(g.sum)}</span>
    </li>`).join('');
}

function renderMerchants(outs) {
  const groups = groupSum(outs, t => t.merchant).slice(0, 8);
  const max = groups.length ? groups[0].sum : 1;
  $('merchantList').innerHTML = groups.map(g => `
    <li>
      <div class="bar-row"><span class="bar-name">${esc(g.key)}</span><span class="bar-sum">${fmt(g.sum)}</span></div>
      <div class="bar-track"><div class="bar-fill" data-w="${(g.sum / max) * 100}"></div></div>
    </li>`).join('');
  requestAnimationFrame(() => {
    document.querySelectorAll('#merchantList .bar-fill').forEach(el => { el.style.width = el.dataset.w + '%'; });
  });
}

function renderDays(outs) {
  const groups = groupSum(outs, t => t.date).sort((a, b) => (a.key < b.key ? -1 : 1));
  const max = groups.reduce((m, g) => Math.max(m, g.sum), 1);
  $('dayChart').innerHTML = groups.map(g => `
    <button type="button" class="day-col${g.key === selectedDay ? ' is-active' : ''}"
            data-day="${g.key}" aria-pressed="${g.key === selectedDay}"
            title="${fmtDate(g.key)} — ${fmt(g.sum)}">
      <span class="day-bar" data-h="${(g.sum / max) * 100}"></span>
      <span class="day-label">${g.key.slice(8, 10)}</span>
    </button>`).join('');
  requestAnimationFrame(() => {
    document.querySelectorAll('#dayChart .day-bar').forEach(el => { el.style.height = el.dataset.h + '%'; });
  });
}

// Раскрывает список трат за выбранный день (повторный клик — закрывает).
function selectDay(day) {
  selectedDay = day === selectedDay ? null : day;
  document.querySelectorAll('#dayChart .day-col').forEach(el => {
    const on = el.dataset.day === selectedDay;
    el.classList.toggle('is-active', on);
    el.setAttribute('aria-pressed', String(on));
  });
  renderDayDetail();
}

function renderDayDetail() {
  const box = $('dayDetail');
  if (!selectedDay) { box.classList.add('hidden'); return; }

  const items = VIEW.filter(t => t.date === selectedDay && t.type === 'out')
    .sort((a, b) => b.amount - a.amount);
  const total = items.reduce((s, t) => s + t.amount, 0);

  $('dayDetailDate').textContent = fmtDate(selectedDay) + ' · ' + plural(items.length, 'трата', 'траты', 'трат');
  $('dayDetailSum').textContent = fmt(total);
  $('dayDetailList').innerHTML = items.map(t => `
    <li>
      <span class="day-item-main">
        <span class="dot" style="background:${catColor(t.cat)}"></span>
        <span class="day-item-name">${esc(t.merchant)}</span>
        <span class="day-item-cat">${esc(t.cat)}</span>
      </span>
      <span class="day-item-sum expense">−${fmt(t.amount)}</span>
    </li>`).join('');

  box.classList.remove('hidden');
}

function renderTable() {
  const q = $('search').value.trim().toLowerCase();
  const cat = $('catFilter').value;
  const list = VIEW.filter(t =>
    (cat === 'all' || t.cat === cat) &&
    (!q || t.merchant.toLowerCase().includes(q) || t.desc.toLowerCase().includes(q))
  );

  $('txCount').textContent = plural(list.length, 'операция', 'операции', 'операций');
  $('txBody').innerHTML = list.slice(0, shown).map(t => `
    <tr>
      <td class="c-date">${fmtDate(t.date)}</td>
      <td class="c-name">${esc(t.merchant)}</td>
      <td class="c-cat"><span class="tag"><span class="dot" style="background:${catColor(t.cat)}"></span>${esc(t.cat)}</span></td>
      <td class="c-amount amount right ${t.type === 'in' ? 'income' : 'expense'}">${t.type === 'in' ? '+' : '−'}${fmt(t.amount)}</td>
    </tr>`).join('');

  $('moreBtn').classList.toggle('hidden', list.length <= shown);
}

function esc(s) {
  return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}

/* ---------- события ---------- */

const drop = $('drop');
$('fileInput').addEventListener('change', e => {
  readFiles(e.target.files);
  e.target.value = '';       // чтобы тот же файл можно было выбрать снова
});
$('fileList').addEventListener('click', e => {
  const x = e.target.closest('.file-chip-x');
  if (x) removeFile(x.dataset.file);
});
['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
drop.addEventListener('drop', e => readFiles(e.dataTransfer.files));

$('dayChart').addEventListener('click', e => {
  const col = e.target.closest('.day-col');
  if (col) selectDay(col.dataset.day);
});
$('dayClose').addEventListener('click', () => selectDay(selectedDay));

$('monthSelect').addEventListener('change', render);
$('search').addEventListener('input', () => { shown = 50; renderTable(); });
$('catFilter').addEventListener('change', () => { shown = 50; renderTable(); });
$('moreBtn').addEventListener('click', () => { shown += 50; renderTable(); });
function resetAll() {
  localStorage.removeItem(STORE_KEY);
  TX = [];
  selectedDay = null;
  $('fileInput').value = '';
  $('dashView').classList.add('hidden');
  $('resetBtn').classList.add('hidden');
  $('uploadView').classList.remove('hidden');
}
$('resetBtn').addEventListener('click', resetAll);

/* ---------- тема ---------- */

function applyTheme(theme) {
  document.documentElement.dataset.theme = theme;
  document.querySelector('#themeBtn .theme-ico').textContent = theme === 'dark' ? '🌙' : '☀️';
  const meta = document.querySelector('meta[name="theme-color"]');
  if (meta) meta.content = theme === 'dark' ? '#12101d' : '#f4f3fb';
  localStorage.setItem(THEME_KEY, theme);
}

$('themeBtn').addEventListener('click', () => {
  applyTheme(document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark');
});

applyTheme(
  localStorage.getItem(THEME_KEY) ||
  (window.matchMedia && window.matchMedia('(prefers-color-scheme: light)').matches ? 'light' : 'dark')
);

/* восстановление ранее загруженных выписок */
try {
  const saved = JSON.parse(localStorage.getItem(STORE_KEY) || 'null');
  if (Array.isArray(saved) && saved.length) {
    // выписки, сохранённые до появления uid, получают его сейчас — иначе сломается склейка
    if (saved.some(t => !t.uid)) {
      const seen = new Map();
      saved.forEach(t => {
        const base = `${t.date}|${t.type}|${t.amount}|${t.desc}`;
        const n = (seen.get(base) || 0) + 1;
        seen.set(base, n);
        t.uid = `${base}|${n}`;
      });
    }
    TX = saved;
    showDashboard();
  }
} catch (e) { /* пустой старт */ }
