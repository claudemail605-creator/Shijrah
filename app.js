/* =====================================================================
   Family Tree viewer
   Data order: 1) data/Master_Family_Tree_Database.xlsx (when served over http)
               2) js/data.js (embedded fallback, works when opened as a file)
               3) "Open Excel" button (pick any .xlsx with the same columns)
   ===================================================================== */
(function () {
  'use strict';

  const XLSX_PATH = 'data/Master_Family_Tree_Database.xlsx';
  const SHEET_NAME = 'Family_Data';

  // Layout constants (px)
  const NW = 180, NH = 72, GX = 16, GY = 66, PAD = 90, SG = 26;
  const unitW = p => p.spouse ? NW * 2 + SG : NW;   // person card (+ spouse card)

  // Excel column -> internal field. Headers are matched ignoring case/underscores/spaces.
  const COLS = {
    individualid: 'id', parentid: 'parentId', generation: 'genFile', gender: 'gender',
    fullname: 'name', nickname: 'nick', fathersname: 'father', mothersname: 'mother',
    spousename: 'spouse', spousefathername: 'spouseFather', dob: 'dob',
    placeofbirth: 'pob', qualification: 'qual', dateofdeath: 'dod',
    placeofburial: 'burial', sourcefile: 'source'
  };

  // Columns shown in the table view / detail panel
  const TABLE_COLS = [
    ['id', 'ID'], ['name', 'Full name'], ['nick', 'Nickname'], ['gen', 'Gen'], ['gender', 'Gender'],
    ['father', "Father"], ['mother', "Mother"], ['spouse', 'Spouse'], ['spouseFather', "Spouse's father"],
    ['dob', 'Date of birth'], ['pob', 'Place of birth'], ['qual', 'Qualification'],
    ['dod', 'Date of death'], ['burial', 'Place of burial'], ['source', 'Source file']
  ];

  const $ = (s, r = document) => r.querySelector(s);
  const el = (tag, cls, html) => { const e = document.createElement(tag); if (cls) e.className = cls; if (html != null) e.innerHTML = html; return e; };
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  /* ------------------------------ State ------------------------------ */
  const S = {
    people: [], byId: new Map(), kids: new Map(), roots: [],
    collapsed: new Set(),
    f: { q: '', gens: new Set(), genders: new Set(), status: '', branch: '', pob: '', qual: '', has: '' },
    view: 'tree', selected: null, sourceLabel: '',
    sort: { key: 'gen', dir: 1 },
    scale: 1, tx: 0, ty: 0, bbox: { w: 0, h: 0 }, pos: new Map()
  };

  /* ---------------------------- Data loading ---------------------------- */
  function norm(h) { return String(h || '').toLowerCase().replace(/[^a-z]/g, ''); }

  function fmtDate(v) {
    if (v == null || v === '') return null;
    if (v instanceof Date && !isNaN(v)) {
      const d = new Date(v.getTime() + 12 * 3600 * 1000); // absorb timezone drift
      const p = n => String(n).padStart(2, '0');
      return `${p(d.getUTCDate())}/${p(d.getUTCMonth() + 1)}/${d.getUTCFullYear()}`;
    }
    return String(v).trim() || null;
  }

  function rowsFromWorkbook(wb) {
    const ws = wb.Sheets[SHEET_NAME] || wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json(ws, { defval: null, raw: true, cellDates: true });
  }

  function rowsFromEmbedded() { return window.FAMILY_DATA || []; }

  function normaliseRows(rows) {
    return rows.map(r => {
      const o = {};
      for (const k of Object.keys(r)) {
        const f = COLS[norm(k)];
        if (!f) continue;
        let v = r[k];
        if (f === 'dob' || f === 'dod') v = fmtDate(v);
        else v = (v == null || String(v).trim() === '') ? null : String(v).trim();
        o[f] = v;
      }
      return o;
    }).filter(o => o.id && o.name);
  }

  async function loadData() {
    // 1) Excel file next to the page (needs http, e.g. VS Code Live Server)
    if (window.XLSX && location.protocol.startsWith('http')) {
      try {
        const res = await fetch(XLSX_PATH, { cache: 'no-store' });
        if (!res.ok) throw new Error(res.status);
        const wb = XLSX.read(await res.arrayBuffer(), { type: 'array', cellDates: true });
        S.sourceLabel = 'Excel file';
        return normaliseRows(rowsFromWorkbook(wb));
      } catch (e) { console.warn('Could not read the Excel file, using embedded data.', e); }
    }
    // 2) embedded fallback
    S.sourceLabel = 'embedded data';
    return normaliseRows(rowsFromEmbedded());
  }

  /* ----------------------------- Build model ----------------------------- */
  function build(rows) {
    S.people = rows;
    S.byId = new Map(rows.map(p => [p.id, p]));
    S.kids = new Map(rows.map(p => [p.id, []]));
    S.roots = [];
    for (const p of rows) {
      if (p.parentId && S.byId.has(p.parentId) && p.parentId !== p.id) S.kids.get(p.parentId).push(p);
      else S.roots.push(p);
    }
    // depth from the tree structure (more reliable than the Generation column)
    const walk = (p, d, seen) => {
      if (seen.has(p.id)) return; seen.add(p.id);
      p.gen = d; p.parent = p.parentId && S.byId.get(p.parentId) || null;
      S.kids.get(p.id).forEach(c => walk(c, d + 1, seen));
    };
    const seen = new Set();
    S.roots.forEach(r => walk(r, 0, seen));
    // anything stuck in a parent loop becomes a root
    rows.forEach(p => { if (!seen.has(p.id)) { S.roots.push(p); walk(p, 0, seen); } });
    // text for searching
    rows.forEach(p => {
      p._hay = [p.id, p.name, p.nick, p.spouse, p.dob, p.pob, p.qual, p.dod, p.burial]
        .filter(Boolean).join(' ').toLowerCase();
      p._initials = initials(p.name);
    });
  }

  function initials(name) {
    const w = name.replace(/^(syed|syeda|mohammed|md|mohd)\s+/i, '').split(/\s+/).filter(Boolean);
    return ((w[0] || '?')[0] + (w[1] ? w[1][0] : '')).toUpperCase();
  }
  const gcls = p => /^m/i.test(p.gender || '') ? 'male' : /^f/i.test(p.gender || '') ? 'female' : 'unknown';

  /* ----------------------------- Date helpers ----------------------------- */
  function parseDate(s) {
    if (!s) return null;
    let m = String(s).match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
    if (m) return new Date(Date.UTC(+m[3], +m[2] - 1, +m[1]));
    m = String(s).match(/^(\d{4})$/);
    if (m) return new Date(Date.UTC(+m[1], 0, 1));
    const d = new Date(s);
    return isNaN(d) ? null : d;
  }
  function yearsBetween(a, b) {
    let y = b.getUTCFullYear() - a.getUTCFullYear();
    const m = b.getUTCMonth() - a.getUTCMonth();
    if (m < 0 || (m === 0 && b.getUTCDate() < a.getUTCDate())) y--;
    return y;
  }
  function ageText(p) {
    const b = parseDate(p.dob); if (!b) return null;
    const d = parseDate(p.dod);
    if (p.dod && !d) return null;
    const y = yearsBetween(b, d || new Date());
    return y < 0 ? null : d ? `${y} years (at death)` : `${y} years`;
  }

  /* ------------------------------- Filters ------------------------------- */
  const isDescendant = (p, branchId) => { for (let c = p; c; c = c.parent) if (c.id === branchId) return true; return false; };

  function filterActive() {
    const f = S.f;
    return !!(f.q || f.gens.size || f.genders.size || f.status || f.branch || f.pob || f.qual || f.has);
  }
  function matches(p) {
    const f = S.f;
    if (f.q && !f.q.split(/\s+/).every(t => p._hay.includes(t))) return false;
    if (f.gens.size && !f.gens.has(p.gen)) return false;
    if (f.genders.size && !f.genders.has(gcls(p))) return false;
    if (f.status === 'deceased' && !p.dod) return false;
    if (f.status === 'living' && p.dod) return false;
    if (f.branch && !isDescendant(p, f.branch)) return false;
    if (f.pob && p.pob !== f.pob) return false;
    if (f.qual && p.qual !== f.qual) return false;
    if (f.has === 'dob' && !p.dob) return false;
    if (f.has === 'spouse' && !p.spouse) return false;
    if (f.has === 'children' && !S.kids.get(p.id).length) return false;
    if (f.has === 'nochildren' && S.kids.get(p.id).length) return false;
    return true;
  }

  /* --------------------------- Sidebar controls --------------------------- */
  function uniq(key) { return [...new Set(S.people.map(p => p[key]).filter(Boolean))].sort((a, b) => a.localeCompare(b)); }

  function buildControls() {
    // generation chips
    const gens = [...new Set(S.people.map(p => p.gen))].sort((a, b) => a - b);
    const gc = $('#genChips'); gc.innerHTML = '';
    gens.forEach(g => gc.append(chip('Gen ' + g, () => toggle(S.f.gens, g), () => S.f.gens.has(g))));
    // gender chips
    const gd = $('#genderChips'); gd.innerHTML = '';
    [['male', 'Male'], ['female', 'Female']].forEach(([k, l]) => gd.append(chip(l, () => toggle(S.f.genders, k), () => S.f.genders.has(k))));

    fillSelect($('#fBranch'), [['', 'Whole family']].concat(
      S.people.filter(p => S.kids.get(p.id).length).sort((a, b) => a.gen - b.gen || a.name.localeCompare(b.name))
        .map(p => [p.id, `${'· '.repeat(p.gen)}${p.name}`])));
    fillSelect($('#fPob'), [['', 'Any place']].concat(uniq('pob').map(v => [v, v])));
    fillSelect($('#fQual'), [['', 'Any qualification']].concat(uniq('qual').map(v => [v, v])));
  }
  function fillSelect(sel, opts) { sel.innerHTML = opts.map(([v, l]) => `<option value="${esc(v)}">${esc(l)}</option>`).join(''); }
  function toggle(set, v) { set.has(v) ? set.delete(v) : set.add(v); }
  function chip(label, onToggle, isOn) {
    const b = el('button', 'chip' + (isOn() ? ' on' : ''), esc(label));
    b.onclick = () => { onToggle(); b.classList.toggle('on', isOn()); refresh(); };
    return b;
  }

  function bindControls() {
    let t;
    $('#search').addEventListener('input', e => { clearTimeout(t); t = setTimeout(() => { S.f.q = e.target.value.trim().toLowerCase(); refresh(); }, 120); });
    [['#fStatus', 'status'], ['#fBranch', 'branch'], ['#fPob', 'pob'], ['#fQual', 'qual'], ['#fHas', 'has']]
      .forEach(([sel, key]) => $(sel).addEventListener('change', e => { S.f[key] = e.target.value; if (key === 'branch' && e.target.value) S.collapsed.delete(e.target.value); refresh(true); }));

    $('#viewSwitch').addEventListener('click', e => {
      const b = e.target.closest('button'); if (!b) return;
      S.view = b.dataset.view;
      [...$('#viewSwitch').children].forEach(x => x.classList.toggle('on', x === b));
      $('#treeView').hidden = S.view !== 'tree';
      $('#tableView').hidden = S.view !== 'table';
      refresh(true);
    });
    $('#btnReset').onclick = resetFilters;
    $('#btnFilters').onclick = () => $('#sidebar').classList.toggle('open');
    $('#btnTheme').onclick = () => {
      const r = document.documentElement;
      const next = r.dataset.theme === 'dark' ? 'light' : 'dark';
      r.dataset.theme = next;
      try { localStorage.setItem('ft-theme', next); } catch (e) { /* ignore */ }
    };
    $('#fileInput').addEventListener('change', async e => {
      const file = e.target.files[0]; if (!file) return;
      try {
        const wb = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true });
        const rows = normaliseRows(rowsFromWorkbook(wb));
        if (!rows.length) throw new Error('No rows with Individual_ID and Full_Name were found.');
        S.sourceLabel = file.name; init(rows); toast(`Loaded ${rows.length} people from ${file.name}`);
      } catch (err) { toast('Could not read that file: ' + err.message); }
      e.target.value = '';
    });
    $('#zoomIn').onclick = () => zoomBy(1.25);
    $('#zoomOut').onclick = () => zoomBy(1 / 1.25);
    $('#zoomFit').onclick = fit;
    $('#expandAll').onclick = () => { S.collapsed.clear(); refresh(true); };
    $('#collapseAll').onclick = () => { S.people.forEach(p => { if (S.kids.get(p.id).length) S.collapsed.add(p.id); }); refresh(true); };
    window.addEventListener('keydown', e => { if (e.key === 'Escape') closePanel(); });
    window.addEventListener('resize', () => { if (S.view === 'tree') fit(); });
  }

  function resetFilters() {
    Object.assign(S.f, { q: '', status: '', branch: '', pob: '', qual: '', has: '' });
    S.f.gens.clear(); S.f.genders.clear();
    $('#search').value = '';
    ['#fStatus', '#fBranch', '#fPob', '#fQual', '#fHas'].forEach(s => $(s).value = '');
    document.querySelectorAll('.chip.on').forEach(c => c.classList.remove('on'));
    refresh(true);
  }

  /* ------------------------------- Rendering ------------------------------- */
  // Filters REMOVE everyone who doesn't match. Matching people whose parents were removed are re-attached to their
  // nearest remaining ancestor; with a Branch filter the chosen person becomes the starting point.
  let visible = null;   // Set of ids shown, or null for "everyone"
  let matchSet = new Set();
  let vKids = new Map(), vRoots = [];

  function computeVisibility() {
    matchSet = new Set(S.people.filter(matches).map(p => p.id));
    if (!filterActive()) { visible = null; vKids = S.kids; vRoots = S.roots; return; }
    visible = new Set(matchSet);
    if (S.f.branch) visible.add(S.f.branch);
    vKids = new Map(S.people.map(p => [p.id, []])); vRoots = [];
    S.people.forEach(p => {
      if (!visible.has(p.id)) return;
      let a = p.parent; while (a && !visible.has(a.id)) a = a.parent;
      (a ? vKids.get(a.id) : vRoots).push(p);
    });
  }
  const shownKids = p => vKids.get(p.id) || [];

  function refresh(refit) {
    computeVisibility();
    renderStats();
    if (S.view === 'tree') renderTree(refit); else renderTable();
    renderPanel();
  }

  function renderStats() {
    const n = S.people.length, gens = new Set(S.people.map(p => p.gen)).size;
    const dec = S.people.filter(p => p.dod).length;
    $('#stats').innerHTML =
      `<div class="stat"><b>${n}</b><span>People</span></div><div class="stat"><b>${gens}</b><span>Generations</span></div><div class="stat"><b>${dec}</b><span>Deceased</span></div>`;
    $('#subtitle').textContent = `${n} members · ${gens} generations · from ${S.sourceLabel}`;
    const rc = $('#resultCount');
    rc.textContent = filterActive() ? `Showing ${S.view === 'tree' && visible ? visible.size : matchSet.size} of ${n} people` : `Showing all ${n} people`;
  }

  /* ---- Tree layout (top-down, subtree-width based) ---- */
  function layout() {
    const pos = new Map();
    const roots = vRoots;
    S.minGen = roots.length ? Math.min(...roots.map(r => r.gen)) : 0;
    const width = p => {
      const k = S.collapsed.has(p.id) ? [] : shownKids(p);
      const w = k.length ? Math.max(unitW(p), k.reduce((s, c) => s + width(c), 0) + GX * (k.length - 1)) : unitW(p);
      p._w = w; return w;
    };
    const place = (p, left) => {
      pos.set(p.id, { x: left + p._w / 2 - unitW(p) / 2, y: (p.gen - S.minGen) * (NH + GY), w: unitW(p) });
      const k = S.collapsed.has(p.id) ? [] : shownKids(p);
      if (!k.length) return;
      const total = k.reduce((s, c) => s + c._w, 0) + GX * (k.length - 1);
      let x = left + (p._w - total) / 2;
      k.forEach(c => { place(c, x); x += c._w + GX; });
    };
    let x = 0;
    roots.forEach(r => { width(r); place(r, x); x += r._w + GX * 2; });
    return { pos, roots };
  }

  function renderTree(refit) {
    const world = $('#world');
    const { pos, roots } = layout();
    S.pos = pos;
    $('#emptyMsg').hidden = !(filterActive() && !matchSet.size);
    if (!pos.size) { world.innerHTML = ''; return; }

    let maxX = 0, maxY = 0, maxGen = 0;
    pos.forEach((v, id) => { maxX = Math.max(maxX, v.x + v.w); maxY = Math.max(maxY, v.y + NH); maxGen = Math.max(maxGen, S.byId.get(id).gen); });
    const W = maxX + PAD * 2, H = maxY + PAD * 2;
    S.bbox = { w: W, h: H };
    const ox = PAD, oy = PAD;   // content offset inside the world

    const active = filterActive();
    let html = '';
    // generation bands
    const gensUsed = new Set([...pos.keys()].map(id => S.byId.get(id).gen));
    gensUsed.forEach(g => {
      if (g % 2) return;
      html += `<div class="band" style="top:${oy + (g - S.minGen) * (NH + GY) - GY / 2}px;height:${NH + GY}px;width:${W}px"><span>Generation ${g}</span></div>`;
    });
    // edges
    let edges = '';
    pos.forEach((v, id) => {
      const p = S.byId.get(id);
      if (S.collapsed.has(id)) return;
      shownKids(p).forEach(c => {
        const q = pos.get(c.id); if (!q) return;
        const x1 = ox + v.x + v.w / 2, y1 = oy + v.y + NH, x2 = ox + q.x + q.w / 2, y2 = oy + q.y, my = y1 + (y2 - y1) / 2;
        const dim = '';
        edges += `<path class="edge${dim}" d="M${x1},${y1} C${x1},${my} ${x2},${my} ${x2},${y2}"/>`;
      });
    });
    html += `<svg width="${W}" height="${H}">${edges}</svg>`;
    // nodes
    pos.forEach((v, id) => {
      const p = S.byId.get(id);
      const nk = shownKids(p).length;
      const cls = ['node', gcls(p)];
      if (S.selected === id) cls.push('selected');
      const tg = nk ? `<button class="tg" data-tg="${esc(id)}" style="left:${ox + v.x + v.w / 2}px;top:${oy + v.y + NH - 11}px" title="${S.collapsed.has(id) ? 'Expand' : 'Collapse'}">${S.collapsed.has(id) ? '+ ' + nk : '−'}</button>` : '';
      html += `<div class="${cls.join(' ')}" data-id="${esc(id)}" style="left:${ox + v.x}px;top:${oy + v.y}px;width:${NW}px;height:${NH}px">
        <div class="avatar">${esc(p._initials)}</div>
        <div class="meta"><div class="nm">${esc(p.name)}</div>${p.nick ? `<div class="nk">“${esc(p.nick)}”</div>` : ''}</div>
        ${p.dod ? '<span class="dag" title="Deceased">†</span>' : ''}</div>`;
      if (p.spouse) {
        const sg = gcls(p) === 'male' ? 'female' : gcls(p) === 'female' ? 'male' : 'unknown';
        const sx = ox + v.x + NW;
        html += `<div class="spouse-link" style="left:${sx}px;top:${oy + v.y + NH / 2}px;width:${SG}px"><i>♥</i></div>
          <div class="node spouse ${sg}${S.selected === id ? ' selected' : ''}" data-id="${esc(id)}" data-spouse="1" style="left:${sx + SG}px;top:${oy + v.y}px;width:${NW}px;height:${NH}px">
          <div class="avatar">${esc(initials(p.spouse))}</div>
          <div class="meta"><div class="nm">${esc(p.spouse)}</div><div class="nk">Spouse of ${esc(p.nick || p.name)}</div></div></div>`;
      }
      html += tg;
    });
    world.innerHTML = html;
    world.style.width = W + 'px'; world.style.height = H + 'px';
    if (refit) fit(); else applyTransform();
  }

  /* ---- Pan / zoom ---- */
  function applyTransform() { $('#world').style.transform = `translate(${S.tx}px,${S.ty}px) scale(${S.scale})`; }
  function fit() {
    const st = $('#treeView'); if (!st || !S.bbox.w) return;
    const vw = st.clientWidth, vh = st.clientHeight;
    // Never shrink below a readable size; if the tree is wider than the screen, start centred on the first root and pan.
    const s = Math.min(1, Math.min(vw / S.bbox.w, (vh - 40) / S.bbox.h));
    S.scale = Math.max(.55, s);
    if (S.bbox.w * S.scale <= vw) S.tx = (vw - S.bbox.w * S.scale) / 2;
    else {
      const first = [...S.pos.values()].sort((a, b) => a.y - b.y || a.x - b.x)[0];
      if (vRoots.length > 1) S.tx = 30 - PAD * S.scale;   // several starting points: begin at the left edge
      else S.tx = vw / 2 - (PAD + (first ? first.x + first.w / 2 : 0)) * S.scale;
    }
    S.ty = Math.max(20, (vh - S.bbox.h * S.scale) / 2 - 10);
    if (S.bbox.h * S.scale > vh) S.ty = 20 - PAD * S.scale + 40;
    applyTransform();
  }
  function zoomAt(f, cx, cy) {
    const ns = Math.min(2.2, Math.max(.2, S.scale * f)), r = ns / S.scale;
    S.tx = cx - (cx - S.tx) * r; S.ty = cy - (cy - S.ty) * r; S.scale = ns; applyTransform();
  }
  function zoomBy(f) { const st = $('#treeView'); zoomAt(f, st.clientWidth / 2, st.clientHeight / 2); }

  function bindStage() {
    const st = $('#treeView'), world = $('#world');
    st.addEventListener('wheel', e => {
      e.preventDefault();
      const r = st.getBoundingClientRect();
      zoomAt(e.deltaY < 0 ? 1.1 : 1 / 1.1, e.clientX - r.left, e.clientY - r.top);
    }, { passive: false });

    let drag = null, moved = false;
    st.addEventListener('pointerdown', e => {
      if (e.target.closest('.node, .tg, .spouse-link, .stage-ctrls, .legend')) return;
      drag = { x: e.clientX, y: e.clientY, tx: S.tx, ty: S.ty }; moved = false; st.classList.add('dragging'); st.setPointerCapture(e.pointerId);
    });
    st.addEventListener('pointermove', e => {
      if (!drag) return;
      const dx = e.clientX - drag.x, dy = e.clientY - drag.y;
      if (Math.abs(dx) + Math.abs(dy) > 3) moved = true;
      S.tx = drag.tx + dx; S.ty = drag.ty + dy; applyTransform();
    });
    const end = () => { drag = null; st.classList.remove('dragging'); };
    st.addEventListener('pointerup', e => { if (drag && !moved && !e.target.closest('.node, .tg')) closePanel(); end(); });
    st.addEventListener('pointercancel', end);

    // Cards: click (details), toggle button (collapse), hover (tooltip)
    world.addEventListener('click', e => {
      const tg = e.target.closest('.tg');
      if (tg) {
        const id = tg.dataset.tg, before = S.pos.get(id);
        S.collapsed.has(id) ? S.collapsed.delete(id) : S.collapsed.add(id); hideTip(); refresh(false);
        // keep the clicked person where it was on screen while the layout re-centres around the new branch
        const after = S.pos.get(id);
        if (before && after) { S.tx -= (after.x + after.w / 2 - before.x - before.w / 2) * S.scale; S.ty -= (after.y - before.y) * S.scale; applyTransform(); }
        return;
      }
      const n = e.target.closest('.node'); if (n) select(n.dataset.id);
    });
    world.addEventListener('mouseover', e => {
      const n = e.target.closest('.node'); if (!n || e.target.closest('.tg')) return;
      if (n === tipNode) return; tipNode = n; n.dataset.spouse ? showSpouseTip(S.byId.get(n.dataset.id), n) : showTip(S.byId.get(n.dataset.id), n);
    });
    world.addEventListener('mouseout', e => {
      const n = e.target.closest('.node'); if (!n) return;
      if (e.relatedTarget && n.contains(e.relatedTarget)) return; hideTip();
    });
    world.addEventListener('mousemove', e => { if (tipNode) moveTip(e); });
  }

  /* ------------------------------ Tooltip ------------------------------ */
  let tipNode = null;
  function showTip(p, node) {
    const rows = [
      ['Generation', 'Gen ' + p.gen], ['Spouse', p.spouse], ['Born', [p.dob, p.pob].filter(Boolean).join(' · ')],
      ['Qualification', p.qual], ['Died', p.dod], ['Children', S.kids.get(p.id).length || null]
    ].filter(r => r[1] != null && r[1] !== '');
    const tip = $('#tooltip');
    tip.innerHTML = `<h4>${esc(p.name)}</h4><div class="sub">${p.nick ? '“' + esc(p.nick) + '” · ' : ''}${esc(p.gender || '')}</div>
      <dl>${rows.map(([k, v]) => `<dt>${k}</dt><dd>${esc(v)}</dd>`).join('')}</dl>
      <div class="more">Click for full details</div>`;
    tip.hidden = false;
    const r = node.getBoundingClientRect(); placeTip(r.right + 12, r.top);
    tip._anchor = r;
  }
  function showSpouseTip(p, node) {
    const rows = [['Spouse of', p.name], ["Spouse's father", p.spouseFather], ['Children together', S.kids.get(p.id).length || null]]
      .filter(r => r[1] != null && r[1] !== '');
    const tip = $('#tooltip');
    tip.innerHTML = `<h4>${esc(p.spouse)}</h4><div class="sub">Spouse</div>
      <dl>${rows.map(([k, v]) => `<dt>${k}</dt><dd>${esc(v)}</dd>`).join('')}</dl>
      <div class="more">Click for ${esc(p.nick || p.name)}’s full details</div>`;
    tip.hidden = false;
    const r = node.getBoundingClientRect(); tip._anchor = r; placeTip(r.right + 12, r.top);
  }
  function placeTip(x, y) {
    const tip = $('#tooltip'), w = tip.offsetWidth, h = tip.offsetHeight;
    if (x + w > innerWidth - 10) x = Math.max(10, x - w - ($('#tooltip')._anchor ? $('#tooltip')._anchor.width : 0) - 24);
    if (y + h > innerHeight - 10) y = innerHeight - h - 10;
    tip.style.left = x + 'px'; tip.style.top = Math.max(10, y) + 'px';
  }
  function moveTip() { /* anchored to the card, nothing to do */ }
  function hideTip() { tipNode = null; $('#tooltip').hidden = true; }

  /* ----------------------------- Detail panel ----------------------------- */
  function select(id, opts = {}) {
    S.selected = id; hideTip();
    document.querySelectorAll('.node.selected').forEach(n => n.classList.remove('selected'));
    const n = $(`.node[data-id="${CSS.escape(id)}"]`); if (n) n.classList.add('selected');
    renderPanel();
    if (opts.focus) focusPerson(id);
    if (S.view === 'table') document.querySelectorAll('#table tbody tr').forEach(r => r.classList.toggle('sel', r.dataset.id === id));
  }
  function closePanel() {
    S.selected = null; $('#panel').hidden = true;
    document.querySelectorAll('.node.selected, tr.sel').forEach(n => n.classList.remove('selected', 'sel'));
  }

  function row(k, v) { return v ? `<div class="p-row"><div class="k">${k}</div><div class="v">${esc(v)}</div></div>` : ''; }
  function sec(title, rows) { const r = rows.join(''); return r ? `<div class="p-sec"><h5>${title}</h5>${r}</div>` : ''; }
  function relBtns(list) { return list.map(p => `<button data-goto="${esc(p.id)}">${esc(p.nick ? p.nick + ' – ' : '')}${esc(p.name)}</button>`).join(''); }

  function renderPanel() {
    const pn = $('#panel');
    const p = S.selected && S.byId.get(S.selected);
    if (!p) { pn.hidden = true; return; }
    const kids = S.kids.get(p.id);
    const sibs = p.parent ? S.kids.get(p.parent.id).filter(x => x.id !== p.id) : [];
    const age = ageText(p);
    const personal = [row('Date of birth', p.dob), row('Age', age), row('Place of birth', p.pob), row('Qualification', p.qual)];
    const family = [row("Father's name", p.father), row("Mother's name", p.mother), row('Spouse', p.spouse), row("Spouse's father", p.spouseFather)];
    const passing = [row('Date of death', p.dod), row('Place of burial', p.burial)];
    const record = [row('Individual ID', p.id), row('Parent ID', p.parentId), row('Source file', p.source)];
    const body = sec('Personal', personal) + sec('Family', family) + sec('Passing', passing);

    let rel = '';
    if (p.parent) rel += `<div class="rel-label">Parent</div><div class="rel">${relBtns([p.parent])}</div>`;
    if (sibs.length) rel += `<div class="rel-label">Siblings (${sibs.length})</div><div class="rel">${relBtns(sibs)}</div>`;
    if (kids.length) rel += `<div class="rel-label">Children (${kids.length})</div><div class="rel">${relBtns(kids)}</div>`;

    pn.innerHTML = `
      <div class="p-head">
        <button class="icon-btn p-close" id="pClose" aria-label="Close">✕</button>
        <div class="avatar p-avatar" style="${''}">${esc(p._initials)}</div>
        <h3>${esc(p.name)}</h3>
        ${p.nick ? `<div class="p-nick">“${esc(p.nick)}”</div>` : ''}
        <div class="p-tags"><span class="tag ${gcls(p)}">${esc(p.gender || 'Gender unknown')}</span><span class="tag">Generation ${p.gen}</span>
          ${p.dod ? '<span class="tag gray">Deceased</span>' : ''}${kids.length ? `<span class="tag gray">${kids.length} child${kids.length > 1 ? 'ren' : ''}</span>` : ''}</div>
      </div>
      ${body.trim() ? body : '<div class="p-none">No further details recorded.</div>'}
      ${rel ? `<div class="p-sec"><h5>Relatives</h5>${rel}</div>` : ''}
      ${sec('Record', record)}
      <div style="height:16px"></div>`;
    pn.className = 'panel ' + gcls(p);
    pn.hidden = false;
    pn.scrollTop = 0;
    $('#pClose').onclick = closePanel;
    // colour the avatar in the panel header
    const av = $('.p-avatar', pn); const g = gcls(p);
    av.style.background = `var(--${g}-soft)`; av.style.color = `var(--${g})`;
    pn.querySelectorAll('[data-goto]').forEach(b => b.onclick = () => select(b.dataset.goto, { focus: true }));
  }

  function focusPerson(id) {
    // make sure the person is visible: expand ancestors, drop filters that hide them
    for (let c = S.byId.get(id).parent; c; c = c.parent) S.collapsed.delete(c.id);
    if (visible && !visible.has(id)) { resetFilters(); toast('Filters cleared to show that person'); }
    if (S.view === 'tree') {
      computeVisibility(); renderTree(false);
      const v = S.pos.get(id); if (!v) return;
      const st = $('#treeView');
      S.scale = Math.max(S.scale, .9);
      S.tx = st.clientWidth / 2 - (PAD + v.x + v.w / 2) * S.scale;
      S.ty = st.clientHeight / 2 - (PAD + v.y + NH / 2) * S.scale;
      const w = $('#world'); w.style.transition = 'transform .45s ease'; applyTransform();
      setTimeout(() => w.style.transition = '', 480);
      const n = $(`.node[data-id="${CSS.escape(id)}"]`); if (n) { n.classList.add('selected', 'pulse'); }
    }
  }

  /* -------------------------------- Table -------------------------------- */
  function renderTable() {
    const t = $('#table');
    const key = S.sort.key, dir = S.sort.dir;
    const rows = S.people.slice().sort((a, b) => {
      const av = a[key] == null ? '' : a[key], bv = b[key] == null ? '' : b[key];
      if (key === 'gen') return (av - bv) * dir || a.id.localeCompare(b.id);
      if (key === 'dob' || key === 'dod') { const x = parseDate(av), y = parseDate(bv); return ((x ? +x : 1e15) - (y ? +y : 1e15)) * dir; }
      return String(av).localeCompare(String(bv), undefined, { numeric: true }) * dir;
    });
    const active = filterActive();
    const list = active ? rows.filter(p => matchSet.has(p.id)) : rows;
    t.innerHTML = `<thead><tr>${TABLE_COLS.map(([k, l]) => `<th data-k="${k}">${l}${k === key ? (dir > 0 ? ' ▲' : ' ▼') : ''}</th>`).join('')}</tr></thead>
      <tbody>${list.map(p => `<tr data-id="${esc(p.id)}" class="${active && !matchSet.has(p.id) ? 'nomatch' : ''}${S.selected === p.id ? ' sel' : ''}">
        ${TABLE_COLS.map(([k]) => { const v = p[k]; return v == null || v === '' ? '<td class="dash">—</td>' : `<td title="${esc(v)}">${esc(v)}</td>`; }).join('')}</tr>`).join('')}</tbody>`;
    $('thead', t).onclick = e => { const th = e.target.closest('th'); if (!th) return; const k = th.dataset.k; S.sort = { key: k, dir: S.sort.key === k ? -S.sort.dir : 1 }; renderTable(); };
    $('tbody', t).onclick = e => { const tr = e.target.closest('tr'); if (tr) select(tr.dataset.id); };
  }

  /* --------------------------------- Misc --------------------------------- */
  let toastT;
  function toast(msg) { const t = $('#toast'); t.textContent = msg; t.hidden = false; clearTimeout(toastT); toastT = setTimeout(() => t.hidden = true, 3200); }

  function init(rows) {
    S.selected = null; S.collapsed.clear();
    build(rows);
    buildControls();
    resetFiltersSilently();
    closePanel();
    refresh(true);
  }
  function resetFiltersSilently() {
    Object.assign(S.f, { q: '', status: '', branch: '', pob: '', qual: '', has: '' });
    S.f.gens.clear(); S.f.genders.clear(); $('#search').value = '';
  }

  let booted = false;
  async function boot() {
    try { const t = localStorage.getItem('ft-theme'); if (t) document.documentElement.dataset.theme = t; } catch (e) { /* ignore */ }
    if (!booted) { bindControls(); bindStage(); booted = true; }
    const rows = await loadData();
    if (!rows.length) { $('#subtitle').textContent = 'No data found – use “Open Excel”'; toast('No data found. Click “Open Excel” to load your file.'); return; }
    init(rows);
  }
  boot();
})();
