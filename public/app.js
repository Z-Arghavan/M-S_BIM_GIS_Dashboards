/* ── State ───────────────────────────────────────────────────────────────── */
const state = {
  standards:     [],
  relationships: [],
  filtered:      [],
  committees:    [],
  selectedNum:   null,
  view:          'catalog',
  q:             '',
  body:          'all',
  committee:     'all',
};

let visNetwork = null;
let graphBuilt = false;
let visNodes   = null;
let visEdges   = null;
let graphStdSet = null;

let chartDrillYear    = null;   // null = year overview; number = month drill-down
let timelineGroups    = new Map();
let timelineSortedYears = [];

/* ── Boot ────────────────────────────────────────────────────────────────── */
(async function init() {
  try {
    const [stds, rels] = await Promise.all([
      fetch('data/standards.json').then(r => r.json()),
      fetch('data/relationships.json').then(r => r.json()),
    ]);

    state.standards     = stds;
    state.relationships = rels;

    // Collect unique committees
    const cSet = new Set();
    stds.forEach(s => {
      if (s['CEN Committee']) cSet.add(s['CEN Committee']);
      if (s['ISO Committee']) cSet.add(s['ISO Committee']);
    });
    state.committees = Array.from(cSet).sort();

    populateCommitteeSelect();
    updateHero();
    applyFilters();
    renderCatalog();
    wireEvents();

  } catch (err) {
    document.getElementById('result-count').textContent =
      'Error loading data. Is the server running and the Excel file present?';
    console.error(err);
  }
})();

/* ── Hero ────────────────────────────────────────────────────────────────── */
function updateHero() {
  document.getElementById('hero-std-count').textContent = state.standards.length;
}

/* ── Committee dropdown ──────────────────────────────────────────────────── */
function populateCommitteeSelect() {
  const sel = document.getElementById('committee-filter');
  state.committees.forEach(c => {
    const opt = document.createElement('option');
    opt.value = c;
    opt.textContent = c;
    sel.appendChild(opt);
  });
}

/* ── Filter logic ────────────────────────────────────────────────────────── */
function applyFilters() {
  const ql = state.q.trim().toLowerCase();
  state.filtered = state.standards.filter(s => {
    if (state.body === 'cen' && s['Is CEN?'] !== 'Yes') return false;
    if (state.body === 'iso' && s['Is ISO?'] !== 'Yes') return false;
    if (state.committee !== 'all' &&
        s['CEN Committee'] !== state.committee &&
        s['ISO Committee'] !== state.committee) return false;
    if (!ql) return true;
    return (
      txt(s['Standard Number']).includes(ql) ||
      txt(s['Title']).includes(ql) ||
      txt(s['Scope/Abstract']).includes(ql)
    );
  });
}

/* ── Catalog render ──────────────────────────────────────────────────────── */
function renderCatalog() {
  const grid  = document.getElementById('catalog-grid');
  const count = document.getElementById('result-count');

  count.textContent = `${state.filtered.length} standard${state.filtered.length !== 1 ? 's' : ''}`;

  if (!state.filtered.length) {
    grid.innerHTML = '<p style="color:var(--muted-fg);font-size:.875rem">No standards match your filters.</p>';
    return;
  }

  grid.innerHTML = state.filtered.map(s => {
    const isCEN     = s['Is CEN?'] === 'Yes';
    const committee = s['CEN Committee'] || s['ISO Committee'] || '—';
    const year      = s['Current Year'] || s['First Year'] || '';
    return `
      <button class="std-card" data-num="${h(s['Standard Number'])}">
        <div class="std-card-top">
          <span class="std-number">${h(s['Standard Number'])}</span>
          <span class="badge ${isCEN ? 'badge-primary' : 'badge-secondary'}">${isCEN ? 'CEN' : 'ISO'}</span>
        </div>
        <p class="std-title">${h(s['Title'] || '')}</p>
        <p class="std-scope">${h(s['Scope/Abstract'] || '')}</p>
        <div class="std-footer">
          <span>${h(committee)}</span>
          <span>${h(String(year))}</span>
        </div>
      </button>`;
  }).join('');

  grid.querySelectorAll('.std-card').forEach(card =>
    card.addEventListener('click', () => openDialog(card.dataset.num))
  );
}

/* ── Dialog ──────────────────────────────────────────────────────────────── */
function openDialog(num) {
  const s = state.standards.find(s => s['Standard Number'] === num);
  if (!s) return;

  state.selectedNum = num;

  const outgoing = state.relationships.filter(r => r['Source Standard Number'] === num);
  const incoming  = state.relationships.filter(r => r['Target Current Number']  === num);
  const stdSet    = new Set(state.standards.map(s => s['Standard Number']));

  // Header
  document.getElementById('dialog-num').textContent = s['Standard Number'];
  document.getElementById('dialog-title').textContent = s['Title'] || '(no title)';

  const badges = [];
  if (s['Is CEN?'] === 'Yes') badges.push(`<span class="badge badge-primary">CEN</span>`);
  if (s['Is ISO?'] === 'Yes') badges.push(`<span class="badge badge-secondary">ISO</span>`);
  if (s['CEN Committee']) badges.push(`<span class="badge badge-outline">${h(s['CEN Committee'])}</span>`);
  if (s['ISO Committee'] && s['ISO Committee'] !== s['CEN Committee'])
    badges.push(`<span class="badge badge-outline">${h(s['ISO Committee'])}</span>`);
  document.getElementById('dialog-badges').innerHTML = badges.join(' ');

  // Fields
  document.getElementById('dialog-scope').textContent        = s['Scope/Abstract'] || '—';
  document.getElementById('dialog-first-year').textContent   = String(s['First Year']   ?? '—');
  document.getElementById('dialog-current-year').textContent = String(s['Current Year'] ?? '—');
  document.getElementById('dialog-ics').textContent          = s['ICS']  || '—';
  document.getElementById('dialog-type').textContent         = s['Type'] || '—';

  // Web link
  const linkWrap = document.getElementById('dialog-link-wrap');
  const linkEl   = document.getElementById('dialog-link');
  if (s['Web Links']) {
    linkEl.href = s['Web Links'];
    linkWrap.hidden = false;
  } else {
    linkWrap.hidden = true;
  }

  // Relationships
  const makeChips = items =>
    items.length
      ? items.map(it => {
          const isKnown = stdSet.has(it.id);
          const cls  = isKnown ? 'rel-chip' : 'rel-chip no-detail';
          const label = it.type ? `${h(it.id)} · <span style="color:var(--muted-fg)">${h(it.type)}</span>` : h(it.id);
          return `<button class="${cls}" data-num="${h(it.id)}">${label}</button>`;
        }).join('')
      : '<span class="empty-note">None</span>';

  document.getElementById('label-outgoing').textContent = `References (${outgoing.length})`;
  document.getElementById('dialog-outgoing').innerHTML  = makeChips(
    outgoing.map(r => ({ id: r['Target Current Number'], type: r['Type'] }))
  );

  document.getElementById('label-incoming').textContent = `Referenced by (${incoming.length})`;
  document.getElementById('dialog-incoming').innerHTML  = makeChips(
    incoming.map(r => ({ id: r['Source Standard Number'], type: r['Type'] }))
  );

  // Wire chip clicks
  document.querySelectorAll('.rel-chip:not(.no-detail)').forEach(chip =>
    chip.addEventListener('click', () => openDialog(chip.dataset.num))
  );

  // Show
  document.getElementById('dialog-overlay').hidden = false;
  document.body.style.overflow = 'hidden';
}

function closeDialog() {
  document.getElementById('dialog-overlay').hidden = true;
  document.body.style.overflow = '';
  state.selectedNum = null;
}

/* ── Timeline chart ──────────────────────────────────────────────────────── */
const MONTH_NAMES = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

function buildBarSVG(labels, counts, { barW = 38, gap = 6, clickable = false, fill = 'hsl(198,90%,32%)' } = {}) {
  const maxCount = Math.max(...counts, 1);
  const padL = 36, padR = 12, padTop = 28, padBot = 36, plotH = 160;
  const svgH = plotH + padTop + padBot;
  const svgW = padL + labels.length * (barW + gap) - gap + padR;

  const ticks = [0, 0.25, 0.5, 0.75, 1].map(frac => {
    const val = Math.round(maxCount * frac);
    const y   = padTop + plotH - frac * plotH;
    return `
      <line x1="${padL - 4}" y1="${y}" x2="${svgW - padR}" y2="${y}"
            stroke="hsl(214,32%,91%)" stroke-width="1"/>
      <text x="${padL - 8}" y="${y + 4}" text-anchor="end"
            font-size="9" fill="hsl(215,16%,47%)">${val}</text>`;
  }).join('');

  const bars = labels.map((label, i) => {
    const count = counts[i];
    const barH  = count > 0 ? Math.max(3, (count / maxCount) * plotH) : 0;
    const x     = padL + i * (barW + gap);
    const y     = padTop + plotH - barH;
    const cursor = clickable && count > 0 ? 'pointer' : 'default';
    return `
      <g class="chart-bar-group" data-label="${h(String(label))}"
         style="cursor:${cursor}" opacity="${count === 0 ? 0.3 : 1}">
        <rect x="${x}" y="${y}" width="${barW}" height="${barH}" rx="3"
              fill="${fill}" opacity="0.82"/>
        ${count > 0 ? `<text x="${x + barW / 2}" y="${y - 5}" text-anchor="middle"
              font-size="10" font-weight="600" fill="hsl(215,30%,12%)">${count}</text>` : ''}
        <text x="${x + barW / 2}" y="${padTop + plotH + 16}" text-anchor="middle"
              font-size="10" fill="hsl(215,16%,47%)">${h(String(label))}</text>
      </g>`;
  }).join('');

  return { svg: `
    <svg width="${svgW}" height="${svgH}" viewBox="0 0 ${svgW} ${svgH}"
         style="display:block;min-width:${svgW}px">
      ${ticks}
      <line x1="${padL}" y1="${padTop}" x2="${padL}" y2="${padTop + plotH}"
            stroke="hsl(214,32%,91%)" stroke-width="1"/>
      ${bars}
    </svg>`, svgW };
}

function renderTimelineChart(groups, sortedYears) {
  const wrap = document.getElementById('timeline-chart');
  if (!wrap) return;

  if (chartDrillYear !== null) {
    renderMonthChart(wrap, groups, chartDrillYear);
  } else {
    renderYearChart(wrap, groups, sortedYears);
  }
}

function renderYearChart(wrap, groups, sortedYears) {
  const yearsAsc = sortedYears
    .filter(y => y !== 'Unknown')
    .slice()
    .sort((a, b) => a - b);

  if (!yearsAsc.length) { wrap.innerHTML = ''; return; }

  const { svg } = buildBarSVG(yearsAsc, yearsAsc.map(y => groups.get(y).length), { clickable: true });
  wrap.innerHTML = svg;

  wrap.querySelectorAll('.chart-bar-group').forEach(g =>
    g.addEventListener('click', () => {
      const yr = Number(g.dataset.label);
      if (!groups.has(yr)) return;
      chartDrillYear = yr;
      renderMonthChart(wrap, groups, yr);
      const target = document.querySelector(`.timeline-group[data-year="${yr}"]`);
      if (target) target.scrollIntoView({ behavior: 'smooth', block: 'start' });
    })
  );
}

function renderMonthChart(wrap, groups, year) {
  const stds = groups.get(year) || [];
  const monthlyCounts = new Array(12).fill(0);
  stds.forEach(s => {
    const { month } = parseYearMonth(s['Current Year'] || s['First Year']);
    if (month >= 1 && month <= 12) monthlyCounts[month - 1]++;
  });

  const { svg } = buildBarSVG(MONTH_NAMES, monthlyCounts, {
    barW: 38, clickable: false, fill: 'hsl(158,64%,38%)'
  });

  wrap.innerHTML = `
    <div class="chart-drilldown-header">
      <button class="chart-back-btn">← All years</button>
      <span class="chart-drilldown-title">${year} — standards by month of publication</span>
    </div>
    ${svg}`;

  wrap.querySelector('.chart-back-btn').addEventListener('click', () => {
    chartDrillYear = null;
    renderYearChart(wrap, timelineGroups, timelineSortedYears);
  });
}

/* ── Timeline render ─────────────────────────────────────────────────────── */
function renderTimeline() {
  const container = document.getElementById('timeline-container');

  if (!state.filtered.length) {
    document.getElementById('timeline-chart').innerHTML = '';
    container.innerHTML = '<p style="color:var(--muted-fg);font-size:.875rem">No standards match your filters.</p>';
    return;
  }

  // Group by integer year (parse YYYY.MM format)
  const groups = new Map();
  state.filtered.forEach(s => {
    const { year } = parseYearMonth(s['Current Year'] || s['First Year']);
    const key = year ?? 'Unknown';
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(s);
  });

  // Sort years descending (newest first), 'Unknown' at end
  const sortedYears = Array.from(groups.keys()).sort((a, b) => {
    if (a === 'Unknown') return 1;
    if (b === 'Unknown') return -1;
    return b - a;
  });

  // Save for chart drill-down re-renders
  timelineGroups     = groups;
  timelineSortedYears = sortedYears;
  chartDrillYear     = null;

  renderTimelineChart(groups, sortedYears);

  container.innerHTML = sortedYears.map(year => {
    // Sort within year by month, then standard number
    const stds = groups.get(year)
      .slice()
      .sort((a, b) => {
        const ma = parseYearMonth(a['Current Year'] || a['First Year']).month || 0;
        const mb = parseYearMonth(b['Current Year'] || b['First Year']).month || 0;
        if (ma !== mb) return ma - mb;
        return (a['Standard Number'] || '').localeCompare(b['Standard Number'] || '');
      });

    const cards = stds.map(s => {
      const isCEN     = s['Is CEN?'] === 'Yes';
      const committee = s['CEN Committee'] || s['ISO Committee'] || '—';
      const { month } = parseYearMonth(s['Current Year'] || s['First Year']);
      const monthLabel = month ? MONTH_NAMES[month - 1] : '';
      return `
        <button class="std-card" data-num="${h(s['Standard Number'])}">
          <div class="std-card-top">
            <span class="std-number">${h(s['Standard Number'])}</span>
            <span class="badge ${isCEN ? 'badge-primary' : 'badge-secondary'}">${isCEN ? 'CEN' : 'ISO'}</span>
          </div>
          <p class="std-title">${h(s['Title'] || '')}</p>
          <p class="std-scope">${h(s['Scope/Abstract'] || '')}</p>
          <div class="std-footer">
            <span>${h(committee)}</span>
            ${monthLabel ? `<span>${h(monthLabel)}</span>` : ''}
          </div>
        </button>`;
    }).join('');

    return `
      <div class="timeline-group" data-year="${h(String(year))}">
        <div class="timeline-year-header">
          <span class="timeline-year-badge">${h(String(year))}</span>
          <div class="timeline-year-divider"></div>
          <span class="timeline-year-count">${stds.length} standard${stds.length !== 1 ? 's' : ''}</span>
        </div>
        <div class="timeline-cards">${cards}</div>
      </div>`;
  }).join('');

  container.querySelectorAll('.std-card').forEach(card =>
    card.addEventListener('click', () => openDialog(card.dataset.num))
  );
}

/* ── Graph highlight helpers ─────────────────────────────────────────────── */
function highlightGraphNode(nodeId) {
  if (!visNetwork || !visNodes || !visEdges) return;
  const connectedNodeIds = new Set(visNetwork.getConnectedNodes(nodeId));
  const connectedEdgeIds = new Set(visNetwork.getConnectedEdges(nodeId));

  visNodes.update(visNodes.get().map(n => ({
    id:      n.id,
    opacity: (n.id === nodeId || connectedNodeIds.has(n.id)) ? 1 : 0.08,
  })));
  visEdges.update(visEdges.get().map(e => ({
    id:    e.id,
    color: {
      color:     connectedEdgeIds.has(e.id) ? 'rgba(110,140,165,0.9)' : 'rgba(110,140,165,0.03)',
      highlight: 'rgba(110,140,165,0.9)',
    },
  })));
}

function clearGraphHighlight() {
  if (!visNodes || !visEdges || !graphStdSet) return;
  visNodes.update(visNodes.get().map(n => ({
    id:      n.id,
    opacity: 1,
    color: {
      background: graphStdSet.has(n.id) ? 'hsl(198,90%,32%)' : 'hsl(35,92%,55%)',
      border:     graphStdSet.has(n.id) ? 'hsl(198,90%,24%)' : 'hsl(35,92%,42%)',
      highlight:  { background: graphStdSet.has(n.id) ? 'hsl(198,90%,45%)' : 'hsl(35,92%,68%)', border: '#555' },
    },
  })));
  visEdges.update(visEdges.get().map(e => ({
    id:    e.id,
    color: { color: 'rgba(110,140,165,0.28)', highlight: 'rgba(110,140,165,0.8)' },
  })));
}

/* ── Graph (vis-network) ─────────────────────────────────────────────────── */
function buildGraph() {
  if (graphBuilt) return;
  graphBuilt = true;

  const stdSet = new Set(state.standards.map(s => s['Standard Number']));
  const stdMap = new Map(state.standards.map(s => [s['Standard Number'], s]));
  graphStdSet  = stdSet;

  // Collect all node IDs (catalogued + referenced)
  const nodeIds = new Set(state.standards.map(s => s['Standard Number']));
  state.relationships.forEach(r => {
    if (r['Source Standard Number']) nodeIds.add(r['Source Standard Number']);
    if (r['Target Current Number'])  nodeIds.add(r['Target Current Number']);
  });

  const nodes = new vis.DataSet(Array.from(nodeIds).map(id => {
    const std  = stdMap.get(id);
    const isCat = stdSet.has(id);

    // Rich tooltip shown on hover
    const tip = document.createElement('div');
    tip.style.cssText = 'max-width:240px;font-size:12px;line-height:1.5;padding:2px 0';
    const parts = [`<strong>${id}</strong>`];
    if (std?.['Title'])
      parts.push(`<span style="color:#555">${std['Title']}</span>`);
    if (std?.['Current Year'] || std?.['First Year'])
      parts.push(`Published: ${std['Current Year'] || std['First Year']}`);
    if (std?.['CEN Committee'] || std?.['ISO Committee'])
      parts.push(std['CEN Committee'] || std['ISO Committee']);
    tip.innerHTML = parts.join('<br>');

    return {
      id,
      title: tip,
      color: {
        background: isCat ? 'hsl(198,90%,32%)' : 'hsl(35,92%,55%)',
        border:     isCat ? 'hsl(198,90%,24%)' : 'hsl(35,92%,42%)',
        highlight:  { background: isCat ? 'hsl(198,90%,45%)' : 'hsl(35,92%,68%)', border: '#555' },
      },
      size:  isCat ? 7 : 4,
      font:  { size: 0 },
    };
  }));

  const edges = new vis.DataSet(
    state.relationships
      .filter(r => r['Source Standard Number'] && r['Target Current Number'])
      .map((r, i) => ({
        id:     i,
        from:   r['Source Standard Number'],
        to:     r['Target Current Number'],
        title:  r['Type'] || undefined,
        color:  { color: 'rgba(110,140,165,0.28)', highlight: 'rgba(110,140,165,0.8)' },
        arrows: { to: { enabled: true, scaleFactor: 0.45 } },
      }))
  );

  visNodes = nodes;
  visEdges = edges;

  const options = {
    nodes:   { shape: 'dot', borderWidth: 1, chosen: true },
    edges:   { width: 0.6, smooth: { type: 'continuous' }, selectionWidth: 0 },
    physics: {
      stabilization: { iterations: 200, updateInterval: 25 },
      barnesHut: {
        gravitationalConstant: -2000,
        centralGravity:        0.3,
        springLength:          90,
        damping:               0.4,
      },
    },
    interaction: {
      hover:             true,
      tooltipDelay:      180,
      navigationButtons: false,
      keyboard:          false,
      zoomView:          true,
    },
  };

  visNetwork = new vis.Network(
    document.getElementById('graph-container'),
    { nodes, edges },
    options
  );

  // After stabilisation: hide loader, auto-fit, enable rearrange button
  visNetwork.on('stabilizationIterationsDone', () => {
    const loading = document.getElementById('graph-loading');
    if (loading) loading.style.display = 'none';
    visNetwork.setOptions({ physics: false });
    visNetwork.fit({ animation: { duration: 600, easingFunction: 'easeInOutQuad' } });
    document.getElementById('graph-physics-toggle').disabled = false;
  });

  // Click node → highlight neighbours + open dialog; click background → clear
  visNetwork.on('click', params => {
    if (params.nodes.length > 0) {
      const id = params.nodes[0];
      highlightGraphNode(id);
      if (stdSet.has(id)) openDialog(id);
    } else {
      clearGraphHighlight();
    }
  });

  // Labels appear at zoom ≥ 1.5
  visNetwork.on('zoom', ({ scale }) => {
    const size = scale >= 1.5 ? 9 : 0;
    nodes.update(nodes.get().map(n => ({ id: n.id, font: { size } })));
  });

  // ── Control buttons ───────────────────────────────────────────────────────
  const STEP = 1.35;
  document.getElementById('graph-zoom-in').addEventListener('click', () =>
    visNetwork.moveTo({ scale: visNetwork.getScale() * STEP, animation: true })
  );
  document.getElementById('graph-zoom-out').addEventListener('click', () =>
    visNetwork.moveTo({ scale: visNetwork.getScale() / STEP, animation: true })
  );
  document.getElementById('graph-fit').addEventListener('click', () =>
    visNetwork.fit({ animation: { duration: 400, easingFunction: 'easeInOutQuad' } })
  );

  let physicsOn = false;
  const physBtn = document.getElementById('graph-physics-toggle');
  physBtn.addEventListener('click', () => {
    physicsOn = !physicsOn;
    visNetwork.setOptions({ physics: { enabled: physicsOn } });
    physBtn.classList.toggle('active', physicsOn);
    physBtn.title = physicsOn ? 'Freeze layout' : 'Rearrange layout';
  });
}

/* ── Event wiring ────────────────────────────────────────────────────────── */
function renderCurrentView() {
  if (state.view === 'catalog')  renderCatalog();
  if (state.view === 'timeline') renderTimeline();
}

function wireEvents() {
  // Search
  document.getElementById('search-input').addEventListener('input', e => {
    state.q = e.target.value;
    applyFilters();
    renderCurrentView();
  });

  // Body filter
  document.getElementById('body-filter').addEventListener('change', e => {
    state.body = e.target.value;
    applyFilters();
    renderCurrentView();
  });

  // Committee filter
  document.getElementById('committee-filter').addEventListener('change', e => {
    state.committee = e.target.value;
    applyFilters();
    renderCurrentView();
  });

  // Tabs
  document.querySelectorAll('.tab-btn').forEach(btn =>
    btn.addEventListener('click', () => switchView(btn.dataset.view))
  );

  // Dialog close
  document.getElementById('dialog-close').addEventListener('click', closeDialog);
  document.getElementById('dialog-overlay').addEventListener('click', e => {
    if (e.target === document.getElementById('dialog-overlay')) closeDialog();
  });
  document.addEventListener('keydown', e => { if (e.key === 'Escape') closeDialog(); });
}

function switchView(view) {
  state.view = view;

  document.querySelectorAll('.tab-btn').forEach(b => {
    const isActive = b.dataset.view === view;
    b.classList.toggle('active', isActive);
    b.setAttribute('aria-selected', isActive);
  });

  document.getElementById('panel-catalog').hidden  = (view !== 'catalog');
  document.getElementById('panel-timeline').hidden = (view !== 'timeline');
  document.getElementById('panel-graph').hidden    = (view !== 'graph');

  if (view === 'graph')    buildGraph();
  if (view === 'timeline') renderTimeline();
}

/* ── Utilities ───────────────────────────────────────────────────────────── */
function parseYearMonth(val) {
  if (val == null) return { year: null, month: null };
  const parts = String(val).split('.');
  const year  = parseInt(parts[0], 10);
  const month = parts[1] ? parseInt(parts[1], 10) : null;
  return {
    year:  isNaN(year)  ? null : year,
    month: isNaN(month) ? null : month,
  };
}

function h(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function txt(str) {
  return String(str ?? '').toLowerCase();
}
