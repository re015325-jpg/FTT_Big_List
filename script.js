/** ==== CONFIG ==== */
const FILE_ID = "1hse75kxI9Gcb7g_tc9LOZ5dYVkFgDcNCrLKyq03cnHg";
let   GID     = "0";
const RANGE   = "A1:AA50000";
const MAX_SHOW_COL_LETTER = "S";
const MAX_SHOW_COL_IDX = colLetterToIdx(MAX_SHOW_COL_LETTER);

/* Tabs for all-tabs export */
const TABS = [
  { name: "United States", gid: "0" },
  { name: "Canada", gid: "918890863" },
  { name: "South America + Caribbean", gid: "506975496" },
  { name: "Europe + Nordic", gid: "1560092023" },
  { name: "East Asia + Oceania + Australia", gid: "1966795654" },
  { name: "Middle East", gid: "1604960979" },
];

/* Remember-last-tab + per-tab scroll keys */
const ACTIVE_GID_KEY = "ftt-active-gid";
const SCROLL_KEY_PREFIX = "ftt-scroll:"; // + GID

/* Wider columns */
const WIDE_COLS = ["K","R"];
const WIDE_IDX  = WIDE_COLS.map(colLetterToIdx);

/* Image pref column (V) */
const IMAGE_PREF_COL_LETTER = "V";
const IMAGE_PREF_COL_IDX = colLetterToIdx(IMAGE_PREF_COL_LETTER);

/* Column indexes (A=0): E=4, F=5, G=6, H=7, I=8, J=9, P=15, Q=16 */
const E_COL_IDX = 4, F_COL_IDX = 5, G_COL_IDX = 6, H_COL_IDX = 7, I_COL_IDX = 8, J_COL_IDX = 9, P_COL_IDX = 15, Q_COL_IDX = 16;

let POLL_MS = 120000;
let ROW_CAP = 2000;
const THUMB_PX = 120;

const MAX_CONCURRENT_IMG = 8;
const IMG_TIMEOUT_MS = 3000;
const LAZY_ROOT_MARGIN = '1200px 0px';

/* Column headers A..S */
const COL_HEADERS = [
  "Location","Alternative Location","Domme","Images/Aliases","Website","Link tree/All links",
  "X/BlueSky","Verified?","Experience?","Price range?","Review/notes","Renown",
  "Link to full review","Travels?","email","Reddit","phone","Services Provided","link"
];

/** ==== DOM ==== */
const $  = s => document.querySelector(s);
const $$ = s => Array.from(document.querySelectorAll(s));
const tbody = $("#tbody");
const colg  = $("#colg");
const meta  = $("#meta");
const imgToggle = $("#imgToggle");
const expandAllBtn = $("#expandAll");
const collapseAllBtn = $("#collapseAll");
const sortColSel = $("#sortCol");
const sortDirSel = $("#sortDir");
const applySortBtn = $("#applySort");
const clearSortBtn = $("#clearSort");
const searchBox = $("#searchBox");
const clearSearchBtn = $("#clearSearch");
const exportBtn = $("#exportCsv");
const exportAllBtn = $("#exportAllCsv");
const scrollWrap = $("#scrollWrap");

/** ==== State ==== */
let currentRows = [];
let currentColCount = 0;
let sortSpec = null;
let lastQuery = "";
let imgCandCache = new Map();

/** ==== Concurrency limiter for images ==== */
let inflight = 0;
const pendingStarts = [];
function schedule(startFn){
  if (inflight < MAX_CONCURRENT_IMG){
    inflight++;
    startFn();
  } else {
    pendingStarts.push(startFn);
  }
}
function doneOne(){
  inflight = Math.max(0, inflight - 1);
  if (pendingStarts.length){
    const fn = pendingStarts.shift();
    inflight++;
    fn();
  }
}

/** ==== Lazy image observer ==== */
const lazyImgObserver = new IntersectionObserver((entries)=>{
  entries.forEach(entry=>{
    if (!entry.isIntersecting) return;
    const starter = entry.target._startLoading;
    if (starter) schedule(starter);
    lazyImgObserver.unobserve(entry.target);
  });
}, { root: scrollWrap, rootMargin: LAZY_ROOT_MARGIN, threshold: 0.01 });

/** ==== Utilities ==== */
function colLetterToIdx(letter){ let n=0; for(let i=0;i<letter.length;i++) n=n*26+(letter.charCodeAt(i)-64); return n-1; }
const URL_RE = /(https?:\/\/[^\s)'"<>]+)/i;
const IMG_EXT_RE   = /\.(png|jpe?g|gif|webp|avif|svg)(\?.*)?$/i;
const IMG_HOST_HINT= /(twimg\.com|fbcdn\.net|cdn|images|media|imgur|ggpht\.com|pinimg\.com)/i;

function debounce(fn, ms){ let t; return (...args)=>{ clearTimeout(t); t=setTimeout(()=>fn(...args), ms); }; }

/* CSV helpers */
function csvEscape(v){
  const s = (v ?? "").toString().replace(/\r?\n/g, " ").trim();
  return /[",\n]/.test(s) ? `"${s.replace(/"/g,'""')}"` : s;
}
function downloadCSV(filename, rows){
  const csv = "\uFEFF" + rows.map(r => r.map(csvEscape).join(",")).join("\r\n");
  const blob = new Blob([csv], {type: "text/csv;charset=utf-8;"});
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  setTimeout(()=>{ URL.revokeObjectURL(url); a.remove(); }, 0);
}

/* Linkify plain URLs (with More/Less) */
function linkifyWithMore(text, {alwaysShow=false} = {}){
  const m = text && text.match(URL_RE);
  if (!m) return null;
  const url = m[1];
  const frag = document.createDocumentFragment();
  const a = document.createElement("a");
  a.href = url; a.target = "_blank"; a.rel = "noopener";
  a.textContent = text; a.className = "cell-link";
  frag.appendChild(a);

  if (alwaysShow || url.length > 60) {
    const btn = document.createElement("span");
    btn.className = "moreless";
    btn.textContent = " More";
    const wrap = document.createElement("div");
    wrap.className = "full-url";
    wrap.textContent = url;

    btn.addEventListener("click", ()=>{
      const td = btn.parentElement;
      td.classList.toggle("show-full");
      btn.textContent = td.classList.contains("show-full") ? " Less" : " More";
    });

    frag.appendChild(btn);
    frag.appendChild(wrap);
  }
  return frag;
}

/* Twitter/X & Reddit helpers */
function buildTwitterLinkFromHandle(txt){
  if (!txt) return null;
  const at = txt.trim();
  if (!/^@[\w.]{1,50}$/i.test(at)) return null;
  const handle = at.slice(1);
  const a = document.createElement("a");
  a.href = `https://x.com/${handle}`;
  a.target = "_blank"; a.rel = "noopener";
  a.textContent = at; a.className = "cell-link";
  return a;
}
function parseRedditUsername(raw) {
  if (!raw) return null;
  let s = String(raw).trim();
  s = s.replace(/^@/,'').replace(/^\/?u\//i,'');
  s = s.split(/\s+/)[0];
  if (!/^[A-Za-z0-9_]{2,25}$/.test(s)) return null;
  return s;
}
function buildRedditLink(raw){
  const user = parseRedditUsername(raw);
  if (!user) return null;
  const a = document.createElement("a");
  a.href = `https://www.reddit.com/user/${encodeURIComponent(user)}`;
  a.target = "_blank"; a.rel = "noopener";
  a.textContent = `u/${user}`;
  a.className = "cell-link";
  return a;
}

/* Phone helpers */
function normalizeDigits(str) { return (str || "").replace(/\D+/g, ""); }
function formatUSPhone(digits) {
  if (digits.length === 10) return `(${digits.slice(0,3)}) ${digits.slice(3,6)}-${digits.slice(6)}`;
  if (digits.length === 11 && digits[0] === '1') return `+1 (${digits.slice(1,4)}) ${digits.slice(4,7)}-${digits.slice(7)}`;
  return digits;
}
function buildTelAnchor(raw) {
  const d = normalizeDigits(raw);
  if (!d) return null;
  const a = document.createElement("a");
  a.href = `tel:${d.length===11 && d[0]==='1' ? '+'+d : (d.length===10 ? '+1'+d : d)}`;
  a.target = "_self";
  a.textContent = formatUSPhone(d);
  a.className = "cell-link";
  return a;
}

/* Avatars/favicons + Twitter helpers */
function extractDomain(url) { try { return new URL(url).hostname; } catch { return ""; } }
function isTwitterDomain(host){
  return /(^|\.)x\.com$/.test(host) || /(^|\.)twitter\.com$/.test(host);
}
function unavatarForDomain(domain) { return domain ? `https://unavatar.io/${encodeURIComponent(domain)}?fallback=false` : ""; }
function faviconForDomain(domain) { return domain ? `https://www.google.com/s2/favicons?domain=${encodeURIComponent(domain)}&sz=128` : ""; }
function unavatarForTwitterHandleStrict(txt) {
  const at = (txt || "").trim();
  if (!/^@[\w.]{1,50}$/i.test(at)) return [];
  const h = at.slice(1);
  return [
    `https://unavatar.io/x/${encodeURIComponent(h)}?fallback=false`,
    `https://unavatar.io/twitter/${encodeURIComponent(h)}?fallback=false`
  ];
}
function extractTwitterHandleFromUrlish(txt){
  const m = (txt || "").match(URL_RE);
  if (!m) return "";
  try {
    const u = new URL(m[1]);
    if (!isTwitterDomain(u.hostname)) return "";
    const parts = u.pathname.split('/').filter(Boolean);
    const handle = parts[0] || "";
    if (/^[A-Za-z0-9_]{1,50}$/.test(handle)) return `@${handle}`;
  } catch {}
  return "";
}
function unavatarForUrlishSkippingTwitter(txt){
  const m = txt && txt.match(URL_RE);
  if (!m) return "";
  const host = extractDomain(m[1]);
  if (!host || isTwitterDomain(host)) return ""; // skip to avoid generic X icon
  return unavatarForDomain(host);
}
function faviconForUrlishSkippingTwitter(txt){
  const m = txt && txt.match(URL_RE);
  if (!m) return "";
  const host = extractDomain(m[1]);
  if (!host || isTwitterDomain(host)) return "";
  return faviconForDomain(host);
}

/* Basic helpers */
function firstUrl(text){
  const m = (text || "").match(URL_RE);
  return m ? m[1] : "";
}
function preferredLandingUrl(row){
  // Click-through: E → G (no F to avoid Linktree logos)
  return firstUrl(row[E_COL_IDX]) || firstUrl(row[G_COL_IDX]) || "";
}
function isOnlyColAText(row){
  const colA = (row[0] || "").trim();
  if (!colA) return false;
  for (let i=1;i<row.length;i++) if ((row[i]||"").trim()!=="") return false;
  return true;
}

/* Cache key per row */
function imageCacheKey(row){
  const v = (row[21]||""); // V is index 21
  const g = (row[G_COL_IDX]||"");
  const e = (row[E_COL_IDX]||"");
  // F intentionally not included (we're not using it for images)
  return `ftt-imgcache:${GID}:${v}|${g}|${e}`;
}
function cacheGetGoodSrc(row){
  try { return localStorage.getItem(imageCacheKey(row)) || ""; } catch { return ""; }
}
function cacheSetGoodSrc(row, rawUrl){
  try { localStorage.setItem(imageCacheKey(row), rawUrl); } catch {}
}

/* Twitter profile image candidates (strict, no favicon) */
function twitterAvatarCandidatesFromGCell(gCellText){
  const list = [];
  list.push(...unavatarForTwitterHandleStrict(gCellText));
  const h = extractTwitterHandleFromUrlish(gCellText);
  if (h) list.push(...unavatarForTwitterHandleStrict(h));
  return list;
}

/* Candidates — order: V → G → E (NO F) + final scan. */
function imageCandidates(row){
  const cached = cacheGetGoodSrc(row);
  const out = [];
  if (cached) out.push(cached);

  // 1) V
  if (IMAGE_PREF_COL_IDX >= 0 && IMAGE_PREF_COL_IDX < row.length) {
    const vtxt = row[IMAGE_PREF_COL_IDX] || "";
    const vm = vtxt.match(URL_RE);
    if (vm) {
      const v = vm[1];
      if (IMG_EXT_RE.test(v) || IMG_HOST_HINT.test(v)) out.push(v);
      else {
        const genV = unavatarForUrlishSkippingTwitter(vtxt) || faviconForUrlishSkippingTwitter(vtxt);
        if (genV) out.push(genV);
      }
    }
  }

  // 2) G (strict twitter avatar) and (if non-twitter URL) site avatar
  const gStrict = twitterAvatarCandidatesFromGCell(row[G_COL_IDX]);
  if (gStrict.length) out.push(...gStrict);
  const gGen = unavatarForUrlishSkippingTwitter(row[G_COL_IDX]) || faviconForUrlishSkippingTwitter(row[G_COL_IDX]);
  if (gGen) out.push(gGen);

  // 3) E (site avatar/fav)
  const eGen = unavatarForUrlishSkippingTwitter(row[E_COL_IDX]) || faviconForUrlishSkippingTwitter(row[E_COL_IDX]);
  if (eGen) out.push(eGen);

  // Final: any direct image anywhere in row
  for (const cell of row) {
    if (!cell) continue;
    const m = cell.match(URL_RE);
    if (!m) continue;
    const url = m[1];
    if (IMG_EXT_RE.test(url) || IMG_HOST_HINT.test(url)) out.push(url);
  }

  // Dedupe
  const seen = new Set();
  return out.filter(u => { if (seen.has(u)) return false; seen.add(u); return true; });
}

function proxify(url){
  try {
    const noScheme = url.replace(/^https?:\/\//i,"");
    return `https://images.weserv.nl/?url=${encodeURIComponent(noScheme)}&w=${THUMB_PX}&h=${THUMB_PX}&fit=cover&q=70`;
  } catch { return url; }
}

/** ==== Collapse state per tab (persisted) ==== */
function loadCollapseState(gid){
  try { return JSON.parse(localStorage.getItem(`ftt-collapse:${gid}`)) || {}; }
  catch { return {}; }
}
function saveCollapseState(gid, map){
  try { localStorage.setItem(`ftt-collapse:${gid}`, JSON.stringify(map)); }
  catch {}
}

/** ==== Show Images preference (persisted) ==== */
function loadShowImages(){
  try {
    const v = localStorage.getItem("ftt-show-images");
    return v === null ? true : v === "true";
  } catch { return true; }
}
function saveShowImages(flag){
  try { localStorage.setItem("ftt-show-images", String(flag)); } catch {}
}

/** ==== Sorting helpers (within groups) ==== */
function comparator(colIdx, dir){
  const mul = (dir === "desc") ? -1 : 1;
  return (ai, bi) => {
    const va = (currentRows[ai] && currentRows[ai][colIdx] || "").toString().toLowerCase();
    const vb = (currentRows[bi] && currentRows[bi][colIdx] || "").toString().toLowerCase();
    if (va < vb) return -1 * mul;
    if (va > vb) return  1 * mul;
    return 0;
  };
}

/** ==== GViz fetch ==== */
function gvizURL(fileId, gid, range){
  const u = new URL(`https://docs.google.com/spreadsheets/d/${fileId}/gviz/tq`);
  u.searchParams.set("gid", gid);
  u.searchParams.set("range", range);
  u.searchParams.set("tqx", "out:json");
  u.searchParams.set("_", Date.now());
  return u.toString();
}
async function fetchRectFor(gid){
  const res  = await fetch(gvizURL(FILE_ID, gid, RANGE), { cache: "no-store" });
  const text = await res.text();
  const fixed = JSON.parse(text.replace(/^[^{]+/, "").replace(/[^}]+$/, ""));
  const rows = (fixed.table?.rows || []).map(r =>
    (r.c || []).map(cell => {
      if (!cell) return "";
      const v = (cell.f ?? cell.v);
      return v == null ? "" : String(v);
    })
  );
  const cols = (fixed.table?.cols || []);
  const colCount = Math.max(cols.length, ...rows.map(r => r.length), MAX_SHOW_COL_IDX+1, 0);
  rows.forEach(r => { while (r.length < colCount) r.push(""); });
  return { rows, colCount };
}
async function fetchRect(){ return fetchRectFor(GID); }

/** ==== Column widths ==== */
function buildColgroup(){
  colg.innerHTML = "";
  const widths = [160];
  for (let i=0;i<=MAX_SHOW_COL_IDX;i++){
    let w = 150;
    if (i < 3) w = 170;
    if (WIDE_IDX.includes(i)) w = 280;
    widths.push(w);
  }
  widths.forEach(px => {
    const col = document.createElement("col");
    col.style.width = px + "px";
    colg.appendChild(col);
  });
}

/** ==== Build groups from state headers ==== */
function buildGroups(rows){
  const state = loadCollapseState(GID);
  const groups = [];
  let current = null;
  for (let i=0; i<rows.length; i++){
    const row = rows[i];
    if (isOnlyColAText(row)) {
      const name = (row[0]||"").trim() || "Untitled";
      current = { headerIndex: i, stateName: name, members: [], collapsed: (state[name] !== undefined ? state[name] : true) };
      groups.push(current);
    } else if (current) {
      current.members.push(i);
    }
  }
  return groups;
}
function persistGroups(groups){
  const map = {};
  for (const g of groups) map[g.stateName] = g.collapsed;
  saveCollapseState(GID, map);
}

/** ==== Helpers: global clamp for text cells ==== */
function makeTextCell(td, text){
  const wrapper = document.createElement("div");
  wrapper.className = "clamp-wrap clamp-3";
  wrapper.textContent = text ?? "";
  td.appendChild(wrapper);
  requestAnimationFrame(()=>{
    if (wrapper.scrollHeight > wrapper.clientHeight + 1) {
      const btn = document.createElement("span");
      btn.className = "moreless";
      btn.textContent = " More";
      btn.addEventListener("click", ()=>{
        wrapper.classList.toggle("clamp-3");
        btn.textContent = wrapper.classList.contains("clamp-3") ? " More" : " Less";
      });
      td.appendChild(btn);
    }
  });
}

/** ==== Build image cell on demand (used at render-time AND on expand) ==== */
function buildImageCell(tdImg, row, rIdx){
  if (!imgToggle.checked) return;
  if (tdImg._imgReady) return; // avoid rebuilding
  tdImg._imgReady = true;

  const cands = imgCandCache.get(rIdx) || imageCandidates(row);
  imgCandCache.set(rIdx, cands);

  if (!cands.length) return;

  const sk  = document.createElement("div"); sk.className = "skeleton";
  const img = document.createElement("img"); img.className = "thumb"; img.loading="lazy"; img.decoding="async";

  let i = 0, timeoutId = null;
  function clearTimer(){ if (timeoutId) { clearTimeout(timeoutId); timeoutId = null; } }
  function tryNext(){
    clearTimer();
    if (i >= cands.length) { sk.remove(); tdImg.textContent = ""; doneOne(); return; }
    const raw = cands[i++];
    img._rawSrc = raw;
    const src = proxify(raw);
    img.src = "";
    img.src = src;
    timeoutId = setTimeout(()=>img.onerror && img.onerror(), IMG_TIMEOUT_MS);
  }
  img.onerror = ()=>{ clearTimer(); tryNext(); };
  img.onload  = () => {
    clearTimer();
    if ((img.naturalWidth||0) < 48 || (img.naturalHeight||0) < 48) { tryNext(); return; }
    cacheSetGoodSrc(row, img._rawSrc || "");
    sk.remove();
    doneOne();
  };

  const starter = ()=>{ tryNext(); };
  img._startLoading = starter;
  lazyImgObserver.observe(img);

  const landing = preferredLandingUrl(row);
  if (landing) {
    const a = document.createElement("a");
    a.href = landing; a.target = "_blank"; a.rel = "noopener"; a.title = "Open link";
    a.appendChild(sk); a.appendChild(img);
    tdImg.appendChild(a);
  } else {
    tdImg.appendChild(sk); tdImg.appendChild(img);
  }
}

/** ==== Render (with per-group sort) ==== */
function render({ rows, colCount }){
  currentRows = rows;
  currentColCount = colCount;

  imgCandCache.clear();
  buildColgroup();
  tbody.innerHTML = "";
  document.body.classList.toggle("hide-images", !imgToggle.checked);

  const groups = buildGroups(rows);
  const groupByHeaderIndex = new Map(groups.map(g => [g.headerIndex, g]));
  const sortedMemberLists = new Map();

  if (sortSpec && typeof sortSpec.colIdx === "number") {
    const cmp = comparator(sortSpec.colIdx, sortSpec.dir);
    for (const g of groups) {
      const sorted = [...g.members].sort(cmp);
      sortedMemberLists.set(g.headerIndex, sorted);
    }
  }

  const renderRowByIndex = (rIdx) => {
    const row = rows[rIdx];

    if (isOnlyColAText(row)) {
      const g = groupByHeaderIndex.get(rIdx);
      const tr = document.createElement("tr");
      if (rIdx === 0) tr.classList.add("is-hidden");
      if (rIdx === 1) tr.classList.add("sticky-second");
      tr.classList.add("state-header");
      if (!g.collapsed) tr.classList.add("expanded");

      const td = document.createElement("td");
      td.colSpan = 1 + (MAX_SHOW_COL_IDX + 1);
      const bar = document.createElement("div");
      bar.className = "state-bar";
      const chev = document.createElement("span"); chev.className = "chev";
      const label = document.createElement("span"); label.textContent = g.stateName;
      bar.appendChild(chev); bar.appendChild(label); td.appendChild(bar); tr.appendChild(td);
      tbody.appendChild(tr);

      tr.addEventListener("click", ()=>{
        g.collapsed = !g.collapsed;
        tr.classList.toggle("expanded", !g.collapsed);
        persistGroups(groups);
        const memberRows = tbody.querySelectorAll(`tr[data-group="${g.headerIndex}"]`);
        memberRows.forEach(mtr => {
          if (g.collapsed) {
            mtr.setAttribute("data-collapsed","true");
          } else {
            mtr.removeAttribute("data-collapsed");
            // Build images now that they’re visible
            const idx = parseInt(mtr.getAttribute("data-row-index"), 10);
            const row = currentRows[idx];
            const tdImg = mtr.querySelector("td.imgCol");
            if (tdImg && imgToggle.checked) buildImageCell(tdImg, row, idx);
          }
        });
      });

      const members = sortedMemberLists.has(g.headerIndex) ? sortedMemberLists.get(g.headerIndex) : g.members;
      for (const mi of members) renderDataRow(mi, g);
      return;
    }

    const belongs = groups.find(g => g.members.includes(rIdx));
    if (!belongs) renderDataRow(rIdx, null);
  };

  function renderDataRow(rIdx, groupObj){
    const row = rows[rIdx];
    const tr = document.createElement("tr");
    tr.setAttribute("data-row-index", String(rIdx));
    if (rIdx === 0) tr.classList.add("is-hidden");
    if (rIdx === 1) tr.classList.add("sticky-second");

    if (groupObj) {
      tr.setAttribute("data-group", groupObj.headerIndex);
      if (groupObj.collapsed) tr.setAttribute("data-collapsed","true");
    }

    // Image cell (click opens E→G)
    const tdImg = document.createElement("td");
    tdImg.className = "imgCol";

    const groupIsCollapsed = !!(groupObj && groupObj.collapsed);
    if (imgToggle.checked && !groupIsCollapsed) {
      buildImageCell(tdImg, row, rIdx);
    }
    tr.appendChild(tdImg);

    // Columns A..S only
    for (let c = 0; c <= MAX_SHOW_COL_IDX; c++) {
      const td  = document.createElement("td");
      const txt = row[c] ?? "";
      const lower = (txt||"").toString().trim().toLowerCase();

      if (c === I_COL_IDX) {
        const span = document.createElement("span");
        span.className = "sentiment";
        if (lower === "positive") { span.textContent = "👍"; span.title = "Positive"; }
        else if (lower === "negative") { span.textContent = "👎"; span.title = "Negative"; }
        else if (lower === "neutral")  { span.textContent = "🤷"; span.title = "Neutral";  }
        else { makeTextCell(td, txt); tr.appendChild(td); continue; }
        td.appendChild(span);
      }
      else if (c === G_COL_IDX) {
        // If G is @handle, render as @handle link.
        // If G is a twitter/x URL, render it *as* @handle link.
        const handleLink = buildTwitterLinkFromHandle(txt);
        if (handleLink) {
          td.appendChild(handleLink);
        } else {
          const at = extractTwitterHandleFromUrlish(txt);
          if (at) {
            const a = document.createElement("a");
            a.href = `https://x.com/${at.slice(1)}`;
            a.target = "_blank"; a.rel = "noopener";
            a.textContent = at;
            a.className = "cell-link";
            td.appendChild(a);
          } else {
            const frag = linkifyWithMore(txt);
            if (frag) td.appendChild(frag); else makeTextCell(td, txt);
          }
        }
      }
      else if (c === E_COL_IDX) {
        const frag = linkifyWithMore(txt, { alwaysShow: true });
        if (frag) td.appendChild(frag); else makeTextCell(td, txt);
      }
      else if (c === F_COL_IDX) {
        // Still show F normally (link text), but DO NOT use F for images.
        const frag = linkifyWithMore(txt, { alwaysShow: true });
        if (frag) td.appendChild(frag); else makeTextCell(td, txt);
      }
      else if (c === P_COL_IDX) {
        const rl = buildRedditLink(txt);
        if (rl) td.appendChild(rl);
        else { const frag = linkifyWithMore(txt); if (frag) td.appendChild(frag); else makeTextCell(td, txt); }
      }
      else if (c === Q_COL_IDX) {
        const tel = buildTelAnchor(txt);
        if (tel) td.appendChild(tel);
        else { const frag = linkifyWithMore(txt); if (frag) td.appendChild(frag); else makeTextCell(td, txt); }
      }
      else {
        const frag = linkifyWithMore(txt);
        if (frag) td.appendChild(frag); else makeTextCell(td, txt);
      }
      tr.appendChild(td);
    }

    tbody.appendChild(tr);
  }

  // Walk rows and render (headers + members) in order
  for (let i=0; i<rows.length && i<ROW_CAP; i++){
    const row = rows[i];
    if (isOnlyColAText(row)) {
      renderRowByIndex(i);
      const g = groups.find(x => x.headerIndex === i);
      if (g && g.members.length) {
        i = g.members[g.members.length - 1];
      }
    } else {
      const inGroup = groups.find(g => g.members.includes(i));
      if (!inGroup) renderRowByIndex(i);
    }
  }

  // Expand/Collapse All (persist)
  expandAllBtn.onclick = () => {
    groups.forEach(g => g.collapsed = false);
    persistGroups(groups);
    tbody.querySelectorAll("tr.state-header").forEach(tr => tr.classList.add("expanded"));
    const members = tbody.querySelectorAll('tr[data-group]');
    members.forEach(mtr => {
      mtr.removeAttribute("data-collapsed");
      const idx = parseInt(mtr.getAttribute("data-row-index"), 10);
      const row = currentRows[idx];
      const tdImg = mtr.querySelector("td.imgCol");
      if (tdImg && imgToggle.checked) buildImageCell(tdImg, row, idx);
    });
  };
  collapseAllBtn.onclick = () => {
    groups.forEach(g => g.collapsed = true);
    persistGroups(groups);
    tbody.querySelectorAll("tr.state-header").forEach(tr => tr.classList.remove("expanded"));
    tbody.querySelectorAll('tr[data-group]').forEach(tr => tr.setAttribute("data-collapsed","true"));
  };

  meta.textContent = `src=gviz • gid=${GID} • rows: ${rows.length} (showing ${Math.min(rows.length, ROW_CAP)}) • cols: ${colCount} • ${new Date().toLocaleTimeString()}`;

  // ---- Restore scroll position after render ----
  restoreScrollFor(GID);
}

/** ==== Search (debounced) ==== */
function applySearch(q, fromRender=false){
  lastQuery = (q || "").trim().toLowerCase();

  const headers = tbody.querySelectorAll("tr.state-header");
  if (lastQuery) {
    headers.forEach(h => h.classList.add("expanded"));
    tbody.querySelectorAll("tr[data-group]").forEach(tr => tr.removeAttribute("data-collapsed"));
    const rows = tbody.querySelectorAll("tr");
    let shown = 0;
    rows.forEach(tr => {
      if (tr.classList.contains("state-header")) { tr.classList.remove("search-row-hide"); return; }
      if (tr.classList.contains("is-hidden")) { tr.classList.add("search-row-hide"); return; }
      const text = tr.textContent.toLowerCase();
      const match = text.includes(lastQuery);
      tr.classList.toggle("search-row-hide", !match);
      if (match) shown++;
    });
    meta.textContent = meta.textContent.replace(/• matches: \d+/, "");
    meta.textContent += ` • matches: ${shown}`;
  } else {
    tbody.querySelectorAll("tr").forEach(tr => tr.classList.remove("search-row-hide"));
    if (!fromRender && currentRows.length) render({ rows: currentRows, colCount: currentColCount });
    meta.textContent = meta.textContent.replace(/• matches: \d+/, "");
  }
}
const debouncedSearch = debounce(()=>applySearch(searchBox.value), 150);

/** ==== Export (current view) ==== */
function exportCurrentViewCSV(){
  const header = COL_HEADERS.slice(0, MAX_SHOW_COL_IDX + 1);
  const out = [header];

  const trs = tbody.querySelectorAll('tr[data-row-index]');
  trs.forEach(tr => {
    if (tr.classList.contains('search-row-hide')) return;
    if (tr.hasAttribute('data-collapsed')) return;
    const idx = parseInt(tr.getAttribute('data-row-index'), 10);
    const row = currentRows[idx];
    if (!row || isOnlyColAText(row)) return;
    out.push(row.slice(0, MAX_SHOW_COL_IDX + 1));
  });

  const ts = new Date().toISOString().replace(/[:T]/g,'-').split('.')[0];
  const name = `ftt-big-list_gid-${GID}_${ts}.csv`;
  downloadCSV(name, out);
}

/** ==== Export (all tabs, full data) ==== */
async function exportAllTabsCSV(){
  const header = ["Region", ...COL_HEADERS.slice(0, MAX_SHOW_COL_IDX + 1)];
  const out = [header];

  for (const tab of TABS) {
    try {
      const pack = await fetchRectFor(tab.gid);
      const rows = pack.rows || [];
      for (let i=0; i<rows.length; i++){
        const r = rows[i];
        if (!r) continue;
        if (isOnlyColAText(r)) continue;
        if (i === 0) continue; // often metadata row
        const slice = r.slice(0, MAX_SHOW_COL_IDX + 1);
        out.push([tab.name, ...slice]);
      }
    } catch (e){
      console.error("Export tab failed", tab, e);
    }
  }

  const ts = new Date().toISOString().replace(/[:T]/g,'-').split('.')[0];
  const name = `ftt-big-list_ALL-TABS_${ts}.csv`;
  downloadCSV(name, out);
}

/** ==== Scroll persistence helpers ==== */
function saveScrollFor(gid, top){
  try { localStorage.setItem(SCROLL_KEY_PREFIX + gid, String(top|0)); } catch {}
}
function loadScrollFor(gid){
  try {
    const v = localStorage.getItem(SCROLL_KEY_PREFIX + gid);
    return v ? parseInt(v, 10) || 0 : 0;
  } catch { return 0; }
}
const debouncedSaveScroll = debounce(()=>{
  saveScrollFor(GID, scrollWrap.scrollTop || 0);
}, 120);

/** ==== Load + poll ==== */
async function load(){
  // save current scroll before reloading (so refresh during same tab keeps place)
  saveScrollFor(GID, scrollWrap.scrollTop || 0);

  meta.textContent = "loading…";
  try {
    const pack = await fetchRect();
    render(pack);
    if (lastQuery) applySearch(lastQuery, /*fromRender*/true);
  } catch(e){
    console.error(e);
    meta.textContent = "Error loading GViz (check sharing / range).";
  }
}
let timer=null;
function startPolling(){ if (timer) clearInterval(timer); timer = setInterval(load, POLL_MS); }

/** ==== UI ==== */
// Tab buttons
$$(".button-row .btn").forEach(btn=>{
  btn.addEventListener("click", async ()=>{
    // persist outgoing tab scroll
    saveScrollFor(GID, scrollWrap.scrollTop || 0);

    $$(".button-row .btn").forEach(b=>b.classList.remove("active"));
    btn.classList.add("active");
    GID = btn.dataset.gid;

    // remember last tab
    try { localStorage.setItem(ACTIVE_GID_KEY, GID); } catch {}

    await load(); // render() will call restoreScrollFor(GID)
    if (lastQuery) applySearch(lastQuery);
  });
});

// Track scrolling to save position
scrollWrap.addEventListener("scroll", debouncedSaveScroll);

imgToggle.addEventListener("change", ()=>{
  saveShowImages(imgToggle.checked);
  if (currentRows.length) render({ rows: currentRows, colCount: currentColCount });
});

// Sort controls
function applySort(){
  const colLetter = sortColSel.value;
  if (!colLetter) { sortSpec = null; }
  else {
    sortSpec = { colIdx: colLetterToIdx(colLetter), dir: (sortDirSel.value === "desc" ? "desc" : "asc") };
  }
  if (currentRows.length) {
    // preserve scroll during sort
    const prevTop = scrollWrap.scrollTop || 0;
    render({ rows: currentRows, colCount: currentColCount });
    scrollWrap.scrollTop = prevTop;
  }
}
sortColSel.addEventListener("change", applySort);
sortDirSel.addEventListener("change", applySort);
applySortBtn.addEventListener("click", applySort);
clearSortBtn.addEventListener("click", ()=>{
  sortColSel.value = ""; sortDirSel.value = "asc"; sortSpec = null; applySort();
});

// Search
searchBox.addEventListener("input", debouncedSearch);
clearSearchBtn.addEventListener("click", ()=>{
  searchBox.value = "";
  applySearch("");
  searchBox.focus();
});

// Export
exportBtn.addEventListener("click", exportCurrentViewCSV);
exportAllBtn.addEventListener("click", exportAllTabsCSV);

/** ==== Restore scroll after render ==== */
function restoreScrollFor(gid){
  const targetTop = loadScrollFor(gid);
  requestAnimationFrame(()=>requestAnimationFrame(()=>{
    scrollWrap.scrollTop = targetTop;
  }));
}

/** ==== Boot ==== */
(async function boot(){
  // restore last-tab if available
  try {
    const stored = localStorage.getItem(ACTIVE_GID_KEY);
    if (stored && $(`.button-row .btn[data-gid="${stored}"]`)) {
      GID = stored;
      $$(".button-row .btn").forEach(b=>b.classList.toggle("active", b.dataset.gid === GID));
    }
  } catch {}
  // restore image toggle
  const v = localStorage.getItem("ftt-show-images");
  const showImages = (v === null ? true : v === "true");
  imgToggle.checked = showImages;
  document.body.classList.toggle("hide-images", !showImages);

  await load();   // render + restoreScrollFor(GID)
  startPolling();
})();
