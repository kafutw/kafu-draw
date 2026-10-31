'use strict';

const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => [...el.querySelectorAll(s)];
const MAX_PRIZES = 20;
const MAX_PER_PRIZE = 100;
const STORE_KEY = 'kafu-draw-setup-v1';

const state = {
  src: 'paste',
  pasteRows: [],
  fileTable: null,      // 二維陣列，上傳檔案的原始內容
  fileRows: [],
  threadsRows: [],
  threadsAuthor: '',
  valid: [],            // 篩選後的名單
  stats: null,
  result: null,
};

/* ---------- 小工具 ---------- */

function toast(msg) {
  const t = $('#toast');
  t.textContent = msg;
  t.hidden = false;
  clearTimeout(toast.timer);
  toast.timer = setTimeout(() => { t.hidden = true; }, 2200);
}

function normUser(s) {
  return String(s || '').trim().replace(/^@+/, '').toLowerCase();
}

function randInt(n) {
  // 用 crypto 亂數，拒絕取樣避免偏差
  const max = Math.floor(0x100000000 / n) * n;
  const buf = new Uint32Array(1);
  do { crypto.getRandomValues(buf); } while (buf[0] >= max);
  return buf[0] % n;
}

function parseTime(s) {
  if (!s) return null;
  const str = String(s).trim().replace(/\//g, '-');
  const d = new Date(/\d{4}-\d{1,2}-\d{1,2} \d/.test(str) ? str.replace(' ', 'T') : str);
  return isNaN(d) ? null : d;
}

function fmtDate(d) {
  const p = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}/${p(d.getMonth() + 1)}/${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
}

function esc(s) {
  return String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function displayName(r) {
  return r.name && normUser(r.name) !== normUser(r.user) ? `${r.name}（${r.user}）` : r.user;
}

/* ---------- 名單來源 ---------- */

function parsePaste(text) {
  const rows = [];
  for (const raw of text.split(/\r?\n/)) {
    const line = raw.trim();
    if (!line) continue;
    let user, msg = '';
    if (line.includes('\t')) {
      const cols = line.split('\t');
      user = cols[0];
      msg = cols.slice(1).join(' ');
    } else {
      const m = line.match(/^(.{1,40}?)\s*[：:]\s*(.*)$/);
      if (m) { user = m[1]; msg = m[2]; } else { user = line; }
    }
    user = user.trim().replace(/^@+/, '');
    if (user) rows.push({ user, name: '', text: msg.trim(), time: '' });
  }
  return rows;
}

async function readTable(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  const buf = await file.arrayBuffer();
  if (ext === 'xlsx' || ext === 'xls') {
    const wb = XLSX.read(buf, { type: 'array', cellDates: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
  }
  // Excel 存的 CSV 常是 Big5，UTF-8 解不開就換 Big5
  let text;
  try { text = new TextDecoder('utf-8', { fatal: true }).decode(buf); }
  catch { text = new TextDecoder('big5').decode(buf); }
  text = text.replace(/^﻿/, '');
  return Papa.parse(text, { skipEmptyLines: true }).data;
}

const COL_GUESS = {
  user: /帳號|username|user ?id|user|profile|handle|id$/i,
  name: /名稱|姓名|暱稱|名字|name|author|作者/i,
  text: /留言|內容|評論|comment|text|message|訊息/i,
  time: /時間|日期|time|date|created/i,
};

function setupColMap(table) {
  const head = table[0] || [];
  const width = Math.max(...table.slice(0, 50).map(r => r.length));
  const guessed = {};
  for (const key of Object.keys(COL_GUESS)) {
    guessed[key] = head.findIndex((h, i) => COL_GUESS[key].test(String(h).trim()) && !Object.values(guessed).includes(i));
  }
  const anyHeader = Object.values(guessed).some(i => i >= 0);
  $('#hasHeader').checked = anyHeader;
  if (guessed.user < 0 && guessed.name >= 0) { guessed.user = guessed.name; guessed.name = -1; }
  if (guessed.user < 0) guessed.user = 0;
  if (!anyHeader && width > 1) guessed.text = 1;

  for (const sel of $$('#colMap select')) {
    const key = sel.dataset.col;
    sel.innerHTML = (key === 'user' ? '' : '<option value="-1">（沒有）</option>') +
      Array.from({ length: width }, (_, i) => {
        const label = anyHeader && head[i] ? String(head[i]).slice(0, 20) : `第 ${i + 1} 欄`;
        return `<option value="${i}">${esc(label)}</option>`;
      }).join('');
    sel.value = String(guessed[key] ?? -1);
  }
  $('#colMap').hidden = false;
}

function rowsFromTable() {
  const t = state.fileTable;
  if (!t) return [];
  const col = {};
  for (const sel of $$('#colMap select')) col[sel.dataset.col] = Number(sel.value);
  const body = $('#hasHeader').checked ? t.slice(1) : t;
  const cell = (r, i) => (i >= 0 && r[i] != null ? String(r[i]).trim() : '');
  return body
    .map(r => ({ user: cell(r, col.user).replace(/^@+/, ''), name: cell(r, col.name), text: cell(r, col.text), time: cell(r, col.time) }))
    .filter(r => r.user);
}

function rowsFromThreads(data) {
  const rows = [];
  const walk = list => {
    for (const c of list || []) {
      if (c.username) rows.push({ user: c.username, name: c.full_name || '', text: c.text || '', time: c.time || '' });
      walk(c.children);
    }
  };
  walk(data.replies);
  return rows;
}

function currentRows() {
  if (state.src === 'file') return state.fileRows;
  if (state.src === 'threads') return state.threadsRows;
  return state.pasteRows;
}

/* ---------- 篩選 ---------- */

function readFilters() {
  return {
    dedupe: $('#fDedupe').checked,
    once: $('#fOnce').checked,
    skipEmpty: $('#fEmpty').checked,
    exclude: $('#fExclude').value.split(/[\n,，]/).map(normUser).filter(Boolean),
    keywords: $('#fKeywords').value.split(/[,，]/).map(s => s.trim()).filter(Boolean),
    kwMode: $('input[name=kwMode]:checked').value,
    minTags: Math.max(0, Number($('#fTags').value) || 0),
    deadline: $('#fDeadline').value ? new Date($('#fDeadline').value) : null,
  };
}

function applyFilters(rows, f) {
  const out = [];
  const drop = { exclude: 0, deadline: 0, keyword: 0, tags: 0, empty: 0, dup: 0 };
  const seen = new Set();
  const ex = new Set(f.exclude);
  for (const r of rows) {
    const key = normUser(r.user);
    if (ex.has(key) || (r.name && ex.has(normUser(r.name)))) { drop.exclude++; continue; }
    if (f.deadline) {
      const t = parseTime(r.time);
      if (t && t > f.deadline) { drop.deadline++; continue; }
    }
    if (f.skipEmpty && !r.text) { drop.empty++; continue; }
    if (f.keywords.length) {
      const txt = r.text.toLowerCase();
      const hit = f.keywords.map(k => txt.includes(k.toLowerCase()));
      if (f.kwMode === 'all' ? !hit.every(Boolean) : !hit.some(Boolean)) { drop.keyword++; continue; }
    }
    if (f.minTags > 0) {
      const tags = new Set((r.text.match(/@[^\s@，,、。！!？?]+/g) || []).map(normUser));
      tags.delete(key);
      if (tags.size < f.minTags) { drop.tags++; continue; }
    }
    if (f.dedupe) {
      if (seen.has(key)) { drop.dup++; continue; }
      seen.add(key);
    }
    out.push({ ...r, key });
  }
  return { out, drop };
}

const DROP_LABEL = {
  exclude: '排除帳號', deadline: '超過截止時間', empty: '沒有留言內容',
  keyword: '沒有關鍵字', tags: '標記人數不足', dup: '重複帳號',
};

function filterSummary(f) {
  const s = [];
  if (f.dedupe) s.push('排除重複帳號（每個帳號只算一次機會）');
  if (f.once) s.push('每個帳號只能抽中一次');
  if (f.skipEmpty) s.push('排除沒有留言內容的');
  if (f.exclude.length) s.push(`排除 ${f.exclude.length} 個指定帳號`);
  if (f.keywords.length) s.push(`留言需包含${f.kwMode === 'all' ? '全部' : '任一'}關鍵字：${f.keywords.join('、')}`);
  if (f.minTags) s.push(`至少標記 ${f.minTags} 位朋友`);
  if (f.deadline) s.push(`留言截止：${fmtDate(f.deadline)}`);
  return s;
}

/* ---------- 獎項 ---------- */

function addPrize(name = '', count = 1) {
  const box = $('#prizes');
  if (box.children.length >= MAX_PRIZES) { toast(`最多 ${MAX_PRIZES} 個獎項`); return; }
  const row = document.createElement('div');
  row.className = 'prize-row';
  row.innerHTML = `
    <input class="pname" type="text" maxlength="40" placeholder="獎品名稱">
    <span class="count"><input class="pcount" type="number" min="1" max="${MAX_PER_PRIZE}">名</span>
    <button class="del" type="button" aria-label="刪除這個獎項">×</button>`;
  $('.pname', row).value = name;
  $('.pcount', row).value = count;
  $('.del', row).onclick = () => {
    if (box.children.length > 1) { row.remove(); refresh(); }
  };
  box.appendChild(row);
  renumberPrizes();
}

function renumberPrizes() {
  $$('#prizes .pname').forEach((el, i) => { el.placeholder = `獎品 ${i + 1}，例如：餐券兩張`; });
}

function readPrizes() {
  return $$('#prizes .prize-row').map((row, i) => ({
    name: $('.pname', row).value.trim() || `獎品 ${i + 1}`,
    count: Math.min(MAX_PER_PRIZE, Math.max(1, Math.floor(Number($('.pcount', row).value) || 1))),
  }));
}

/* ---------- 畫面更新 ---------- */

function refresh() {
  const rows = currentRows();
  const f = readFilters();
  const { out, drop } = applyFilters(rows, f);
  state.valid = out;
  state.stats = { raw: rows.length, valid: out.length, people: new Set(out.map(r => r.key)).size, drop };

  const sum = $('#summary');
  if (!rows.length) {
    sum.hidden = true;
    $('#previewWrap').hidden = true;
  } else {
    const drops = Object.entries(drop).filter(([, n]) => n).map(([k, n]) => `<li>${DROP_LABEL[k]}：${n} 則</li>`).join('');
    sum.innerHTML = `共 ${rows.length} 則留言，可以參加抽獎的有 <strong>${out.length}</strong> 則` +
      (f.dedupe ? '' : `（${state.stats.people} 個不同帳號）`) +
      (drops ? `<ul>${drops}</ul>` : '');
    sum.hidden = false;
    $('#previewCount').textContent = out.length;
    $('#previewTable tbody').innerHTML = out.slice(0, 500).map((r, i) =>
      `<tr><td>${i + 1}</td><td>${esc(displayName(r))}</td><td>${esc(r.text)}</td><td>${esc(r.time)}</td></tr>`).join('') +
      (out.length > 500 ? `<tr><td></td><td colspan="3">…只列出前 500 則</td></tr>` : '');
    $('#previewWrap').hidden = false;
  }

  const prizes = readPrizes();
  const need = prizes.reduce((a, p) => a + p.count, 0) + backupCount();
  const maxOne = Math.max(...prizes.map(p => p.count));
  let note, ok = false;
  if (!out.length) note = '先放入參加名單。';
  else if (f.once && state.stats.people < need) note = `得獎名額加備取共 ${need} 位，但只有 ${state.stats.people} 個不同帳號，請減少名額或放寬條件。`;
  else if (!f.once && state.stats.people < maxOne) note = `有獎項要抽 ${maxOne} 位，但只有 ${state.stats.people} 個不同帳號，請減少名額或放寬條件。`;
  else { note = `將從 ${out.length} 則留言抽出 ${need} 位${backupCount() ? `（含備取 ${backupCount()} 位）` : ''}。`; ok = true; }
  $('#goNote').textContent = note;
  $('#drawBtn').disabled = !ok;
  saveSetup();
}

function backupCount() {
  return Math.min(50, Math.max(0, Math.floor(Number($('#backup').value) || 0)));
}

/* ---------- 開獎 ---------- */

function drawWinners(entries, prizes, backup, once) {
  const used = new Set();
  const pickFrom = (pool, n, extraUsed) => {
    const picked = [];
    const local = new Set(extraUsed);
    let cand = pool.filter(r => !local.has(r.key));
    while (picked.length < n && cand.length) {
      const r = cand[randInt(cand.length)];
      picked.push(r);
      local.add(r.key);
      cand = cand.filter(c => c.key !== r.key);
    }
    return picked;
  };
  const results = prizes.map(p => {
    const w = pickFrom(entries, p.count, once ? used : []);
    if (once) w.forEach(r => used.add(r.key));
    return { name: p.name, count: p.count, winners: w };
  });
  const allWon = new Set(results.flatMap(p => p.winners.map(r => r.key)));
  // 備取一律不跟正取重複
  const backups = backup ? pickFrom(entries, backup, allWon) : [];
  return { prizes: results, backups };
}

let skipAnim = false;

async function animateDraw(res) {
  show('drawing');
  skipAnim = false;
  const reel = $('#reel');
  const list = $('#stageList');
  const pool = state.valid;
  const groups = [...res.prizes.map(p => ({ name: p.name, w: p.winners })), ...(res.backups.length ? [{ name: '備取', w: res.backups }] : [])];
  const total = groups.reduce((a, g) => a + g.w.length, 0);
  const per = Math.max(250, Math.min(1400, 9000 / Math.max(1, total)));
  const sleep = ms => new Promise(r => setTimeout(r, skipAnim ? 0 : ms));

  for (const g of groups) {
    $('#stagePrize').textContent = g.name;
    list.innerHTML = '';
    for (const w of g.w) {
      if (!skipAnim) {
        const end = performance.now() + per;
        while (performance.now() < end && !skipAnim) {
          reel.textContent = pool[randInt(pool.length)].user;
          await new Promise(r => requestAnimationFrame(r));
        }
      }
      reel.textContent = w.user;
      const li = document.createElement('li');
      li.textContent = w.user;
      list.appendChild(li);
      await sleep(Math.min(500, per / 2));
    }
    await sleep(700);
  }
}

async function runDraw() {
  const f = readFilters();
  const prizes = readPrizes();
  const res = drawWinners(state.valid, prizes, backupCount(), f.once);
  const pack = r => [r.user, r.name && normUser(r.name) !== normUser(r.user) ? r.name : ''];
  state.result = {
    v: 1,
    t: $('#title').value.trim() || '留言抽獎',
    u: $('#postUrl').value.trim(),
    d: Date.now(),
    n: [state.stats.raw, state.stats.valid, state.stats.people],
    f: filterSummary(f),
    p: res.prizes.map(p => ({ n: p.name, w: p.winners.map(pack) })),
    b: res.backups.map(pack),
  };
  // 下載 CSV 用，含留言內容，不放進網址
  state.detail = res;
  await animateDraw(res);
  history.replaceState(null, '', resultUrl(state.result));
  markOwned(state.result.d);
  renderResult(state.result, true);
}

function resultUrl(r) {
  return location.origin + location.pathname + '#r=' + LZString.compressToEncodedURIComponent(JSON.stringify(r));
}

/* ---------- 開獎結果 ---------- */

function renderResult(r, owner) {
  document.title = `${r.t}｜開獎結果`;
  $('#rTitle').textContent = r.t;
  const [raw, valid, people] = r.n;
  $('#rMeta').textContent = `開獎時間 ${fmtDate(new Date(r.d))}・有效名單 ${valid} 則（共 ${raw} 則留言、${people} 個帳號）`;
  const post = $('#rPost');
  post.innerHTML = '';
  if (/^https?:\/\//.test(r.u)) {
    const a = document.createElement('a');
    a.href = r.u; a.target = '_blank'; a.rel = 'noopener';
    a.textContent = '看活動貼文';
    post.appendChild(a);
  }
  const card = (title, list, cls) => `
    <section class="card prize-card ${cls}">
      <h3>${esc(title)}<small>${list.length} 名</small></h3>
      <ol class="winners">${list.map(([u, n], i) =>
        `<li><span class="no">${i + 1}</span><span class="who">${esc(u)}${n ? `<span class="sub">${esc(n)}</span>` : ''}</span></li>`).join('')}</ol>
    </section>`;
  $('#rPrizes').innerHTML = r.p.map(p => card(p.n, p.w, '')).join('') + (r.b.length ? card('備取（依序遞補）', r.b, 'backup') : '');
  $('#rFilters').innerHTML = r.f.length ? `篩選條件<ul>${r.f.map(s => `<li>${esc(s)}</li>`).join('')}</ul>` : '沒有設定篩選條件';
  $('#ownerActions').hidden = !owner;
  $('#publicFoot').hidden = owner;
  show('result');
  window.scrollTo(0, 0);
}

function postText(r) {
  const lines = [`🎉【${r.t}】開獎結果`, ''];
  for (const p of r.p) {
    lines.push(`🏆 ${p.n}（${p.w.length} 名）`);
    lines.push(p.w.map(([u]) => '@' + u).join('、'), '');
  }
  if (r.b.length) lines.push(`備取：${r.b.map(([u]) => '@' + u).join('、')}`, '');
  lines.push('恭喜得獎的朋友，請留意私訊通知喔！', '', `完整開獎結果：${resultUrl(r)}`);
  return lines.join('\n');
}

function downloadCsv() {
  const r = state.result;
  const rows = [['獎項', '順位', '帳號', '顯示名稱', '留言', '留言時間']];
  const add = (prize, list) => list.forEach((w, i) => rows.push([prize, i + 1, w.user, w.name, w.text, w.time]));
  if (state.detail) {
    state.detail.prizes.forEach(p => add(p.name, p.winners));
    add('備取', state.detail.backups);
  } else {
    r.p.forEach(p => p.w.forEach(([u, n], i) => rows.push([p.n, i + 1, u, n, '', ''])));
    r.b.forEach(([u, n], i) => rows.push(['備取', i + 1, u, n, '', '']));
  }
  const csv = '﻿' + rows.map(row => row.map(c => `"${String(c ?? '').replace(/"/g, '""')}"`).join(',')).join('\r\n');
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
  a.download = `${r.t}-得獎名單.csv`;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}

async function copy(text, msg) {
  try { await navigator.clipboard.writeText(text); toast(msg); }
  catch { prompt('請手動複製：', text); }
}

/* ---------- 開獎圖（1080×1350，人少放大字、太多人自動分頁） ---------- */

async function makeImages(r) {
  try { await Promise.all([document.fonts.load('900 40px "Noto Sans TC"'), document.fonts.load('500 30px "Noto Sans TC"')]); } catch {}
  // 得獎人少就放大字，一頁放得下的最大倍率優先，放不下才分頁
  for (const scale of [1.7, 1.4, 1.2]) {
    const pages = layoutImages(r, scale, true);
    if (pages === 1) return layoutImages(r, scale, false);
  }
  return layoutImages(r, 1, false);
}

function layoutImages(r, scale, dry) {
  const W = 1080, H = 1350, PAD = 72, BOTTOM = H - 90;
  const FONT = '"Noto Sans TC", "PingFang TC", sans-serif';
  const brand = '#347E86', gold = '#c8962e';
  const groups = [...r.p.map(p => ({ title: p.n, list: p.w, backup: false })), ...(r.b.length ? [{ title: '備取', list: r.b, backup: true }] : [])];
  const measure = document.createElement('canvas').getContext('2d');
  const pages = [];
  let ctx, y;
  const newPage = first => {
    if (dry) { ctx = measure; pages.push(null); }
    else {
      const c = document.createElement('canvas');
      c.width = W; c.height = H;
      ctx = c.getContext('2d');
      pages.push(c);
      ctx.fillStyle = '#f7f5ef'; ctx.fillRect(0, 0, W, H);
      ctx.fillStyle = brand; ctx.fillRect(0, 0, W, first ? 300 : 150);
      ctx.fillStyle = '#fff'; ctx.textBaseline = 'top';
      ctx.font = `700 34px ${FONT}`;
      ctx.fillText('開獎結果', PAD, first ? 64 : 52);
      if (first) {
        ctx.font = `900 60px ${FONT}`;
        wrapText(ctx, r.t, PAD, 118, W - PAD * 2, 74, 2);
        ctx.font = `400 26px ${FONT}`;
        ctx.fillStyle = 'rgba(255,255,255,.85)';
        ctx.fillText(`${fmtDate(new Date(r.d))}・有效名單 ${r.n[1]} 則`, PAD, 252);
      }
    }
    y = (first ? 300 : 150) + 56;
  };
  newPage(true);

  for (const g of groups) {
    const cols = g.list.length > 8 ? 3 : g.list.length > 3 ? 2 : 1;
    const nameSize = Math.round((cols === 1 ? 40 : cols === 2 ? 34 : 28) * scale);
    const rowH = Math.round(nameSize * 1.6);
    const headH = Math.round(34 * scale) + 28;
    const colW = (W - PAD * 2) / cols;
    let i = 0, cont = false;
    while (i < g.list.length) {
      if (y + headH + rowH > BOTTOM) newPage(false);
      if (!dry) {
        ctx.fillStyle = g.backup ? '#7a858b' : gold;
        ctx.font = `700 ${Math.round(34 * scale)}px ${FONT}`;
        ctx.fillText(`${g.backup ? '' : '🏆 '}${g.title}${cont ? '（續）' : ''}`, PAD, y);
      }
      y += headH;
      const room = Math.max(1, Math.floor((BOTTOM - y) / rowH)) * cols;
      const chunk = g.list.slice(i, i + room);
      if (!dry) chunk.forEach(([u], k) => {
        const cx = PAD + (k % cols) * colW;
        const cy = y + Math.floor(k / cols) * rowH;
        ctx.fillStyle = '#9aa5ab';
        ctx.font = `500 ${Math.round(nameSize * .6)}px ${FONT}`;
        ctx.fillText(String(i + k + 1), cx, cy + nameSize * .3);
        ctx.fillStyle = '#1f2a2c';
        ctx.font = `500 ${nameSize}px ${FONT}`;
        fitText(ctx, u, cx + nameSize * 1.1, cy, colW - nameSize * 1.1 - 16);
      });
      y += Math.ceil(chunk.length / cols) * rowH + 40;
      i += chunk.length;
      cont = true;
    }
  }
  if (dry) return pages.length;
  pages.forEach((c, k) => {
    const g = c.getContext('2d');
    g.fillStyle = '#9aa5ab'; g.font = `400 24px ${FONT}`; g.textBaseline = 'alphabetic';
    g.fillText(pages.length > 1 ? `第 ${k + 1}／${pages.length} 頁` : '恭喜得獎的朋友！', PAD, H - 44);
  });
  return pages;
}

function wrapText(ctx, text, x, y, maxW, lineH, maxLines) {
  let line = '', n = 0;
  for (const ch of text) {
    if (ctx.measureText(line + ch).width > maxW && line) {
      if (++n >= maxLines) { fitText(ctx, line + '…', x, y, maxW); return; }
      ctx.fillText(line, x, y); y += lineH; line = ch;
    } else line += ch;
  }
  ctx.fillText(line, x, y);
}

function fitText(ctx, text, x, y, maxW) {
  let t = text;
  while (ctx.measureText(t).width > maxW && t.length > 1) t = t.slice(0, -2) + '…';
  ctx.fillText(t, x, y);
}

async function showImages() {
  const pages = await makeImages(state.result);
  const list = $('#imgList');
  list.innerHTML = '';
  pages.forEach((c, k) => {
    const url = c.toDataURL('image/png');
    const fig = document.createElement('figure');
    fig.innerHTML = `<img alt="開獎圖第 ${k + 1} 頁" src="${url}"><figcaption><a download="${esc(state.result.t)}-開獎圖${pages.length > 1 ? '-' + (k + 1) : ''}.png" href="${url}">下載這張</a></figcaption>`;
    list.appendChild(fig);
  });
  $('#imgModal').hidden = false;
}

/* ---------- 頁面切換與暫存 ---------- */

// 自己開過的獎，重新整理後仍然看得到操作按鈕
function markOwned(d) {
  try {
    const own = JSON.parse(localStorage.getItem(STORE_KEY + '-own') || '[]');
    localStorage.setItem(STORE_KEY + '-own', JSON.stringify([d, ...own].slice(0, 30)));
  } catch {}
}

function isOwned(d) {
  try { return JSON.parse(localStorage.getItem(STORE_KEY + '-own') || '[]').includes(d); } catch { return false; }
}

function show(id) {
  for (const m of ['setup', 'drawing', 'result']) $('#' + m).hidden = m !== id;
}

function saveSetup() {
  try {
    localStorage.setItem(STORE_KEY, JSON.stringify({
      title: $('#title').value, postUrl: $('#postUrl').value, paste: $('#pasteBox').value,
      prizes: readPrizes(), backup: $('#backup').value,
      f: { dedupe: $('#fDedupe').checked, once: $('#fOnce').checked, empty: $('#fEmpty').checked,
           exclude: $('#fExclude').value, keywords: $('#fKeywords').value,
           kwMode: $('input[name=kwMode]:checked').value, tags: $('#fTags').value, deadline: $('#fDeadline').value },
    }));
  } catch {}
}

function loadSetup() {
  let s = null;
  try { s = JSON.parse(localStorage.getItem(STORE_KEY)); } catch {}
  if (!s) { addPrize(); return; }
  $('#title').value = s.title || '';
  $('#postUrl').value = s.postUrl || '';
  $('#pasteBox').value = s.paste || '';
  state.pasteRows = parsePaste($('#pasteBox').value);
  (s.prizes?.length ? s.prizes : [{ name: '', count: 1 }]).forEach(p => addPrize(/^獎品 \d+$/.test(p.name) ? '' : p.name, p.count));
  $('#backup').value = s.backup || 0;
  const f = s.f || {};
  $('#fDedupe').checked = f.dedupe !== false;
  $('#fOnce').checked = f.once !== false;
  $('#fEmpty').checked = !!f.empty;
  $('#fExclude').value = f.exclude || '';
  $('#fKeywords').value = f.keywords || '';
  const radio = $(`input[name=kwMode][value="${f.kwMode || 'any'}"]`);
  if (radio) radio.checked = true;
  $('#fTags').value = f.tags || 0;
  $('#fDeadline').value = f.deadline || '';
}

function setSource(src) {
  state.src = src;
  $$('.tabs button').forEach(b => b.setAttribute('aria-selected', String(b.dataset.src === src)));
  $$('.pane').forEach(p => { p.hidden = p.dataset.pane !== src; });
  refresh();
}

function bindDrop(input, onFile) {
  const box = input.closest('.drop');
  input.addEventListener('change', () => input.files[0] && onFile(input.files[0]));
  box.addEventListener('dragover', e => { e.preventDefault(); box.classList.add('over'); });
  box.addEventListener('dragleave', () => box.classList.remove('over'));
  box.addEventListener('drop', e => {
    e.preventDefault(); box.classList.remove('over');
    if (e.dataTransfer.files[0]) onFile(e.dataTransfer.files[0]);
  });
}

function init() {
  loadSetup();

  $$('.tabs button').forEach(b => b.addEventListener('click', () => setSource(b.dataset.src)));
  $('#pasteBox').addEventListener('input', () => { state.pasteRows = parsePaste($('#pasteBox').value); refresh(); });

  bindDrop($('#fileInput'), async file => {
    try {
      state.fileTable = await readTable(file);
      if (!state.fileTable.length) throw new Error('empty');
      $('#fileLabel').textContent = `${file.name}（${state.fileTable.length} 列）`;
      setupColMap(state.fileTable);
      state.fileRows = rowsFromTable();
      refresh();
    } catch {
      toast('這個檔案讀不出來，換成 CSV 再試試');
    }
  });
  $('#colMap').addEventListener('change', () => { state.fileRows = rowsFromTable(); refresh(); });

  bindDrop($('#threadsInput'), async file => {
    try {
      const data = JSON.parse(await file.text());
      if (!Array.isArray(data.replies)) throw new Error('format');
      state.threadsRows = rowsFromThreads(data);
      const meta = data.meta || {};
      $('#threadsLabel').textContent = `${file.name}（${state.threadsRows.length} 則留言）`;
      if (meta.url && !$('#postUrl').value) $('#postUrl').value = meta.url;
      if (meta.author) {
        const ex = $('#fExclude');
        const have = ex.value.split(/\n/).map(normUser);
        if (!have.includes(normUser(meta.author))) ex.value = (ex.value.trim() ? ex.value.trim() + '\n' : '') + meta.author;
      }
      const note = $('#threadsNote');
      note.hidden = !meta.note;
      note.textContent = meta.note ? `提醒：${meta.note}` : '';
      refresh();
    } catch {
      toast('這不是 threads-dump 的 raw.json');
    }
  });

  $('#addPrize').addEventListener('click', () => { addPrize(); refresh(); });
  $('#setup').addEventListener('input', e => { if (!e.target.closest('#pasteBox, .drop')) refresh(); });
  $('#setup').addEventListener('change', e => { if (e.target.matches('input[type=checkbox], input[type=radio]')) refresh(); });

  $('#drawBtn').addEventListener('click', runDraw);
  $('#skipBtn').addEventListener('click', () => { skipAnim = true; });
  $('#backToSetup').addEventListener('click', () => {
    if (!confirm('回到設定頁之後，這次的開獎結果網址還是有效，但畫面會離開。確定嗎？')) return;
    history.replaceState(null, '', location.pathname);
    show('setup');
  });
  $('#copyUrl').addEventListener('click', () => copy(resultUrl(state.result), '已複製開獎結果網址'));
  $('#copyPost').addEventListener('click', () => copy(postText(state.result), '已複製貼文文字'));
  $('#dlCsv').addEventListener('click', downloadCsv);
  $('#makeImg').addEventListener('click', showImages);
  $('#closeModal').addEventListener('click', () => { $('#imgModal').hidden = true; });
  $('#imgModal').addEventListener('click', e => { if (e.target.id === 'imgModal') $('#imgModal').hidden = true; });

  refresh();

  const m = location.hash.match(/^#r=(.+)$/);
  if (m) {
    try {
      const r = JSON.parse(LZString.decompressFromEncodedURIComponent(m[1]));
      if (r && r.v === 1) { state.result = r; renderResult(r, isOwned(r.d)); return; }
    } catch {}
    toast('開獎結果網址不完整，請確認有整串複製');
  }
  show('setup');
}

init();
