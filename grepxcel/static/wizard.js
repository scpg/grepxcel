// ── State ────────────────────────────────────────────────────────────────
let selectedRef = null;
let selectedAction = null;
let cellData = {};   // ref → {value, choice, name, …}
let sheetData = null;
let currentStats = {};
let csvText = '';

// ── AG Grid state ─────────────────────────────────────────────────────────
let gridApi = null;
let _inRangeRefs = new Set();

// ── Alpine store ──────────────────────────────────────────────────────────
document.addEventListener('alpine:init', () => {
  Alpine.store('gx', {
    activeTab:      'classify',  // 'classify' | 'extract' | 'config' | 'logs'
    panelMode:      'hint',      // 'hint' | 'classify' | 'table'
    selectedAction: null,        // 'L' | 'V' | 'T' | 'I' | null
    trmOpen:        false,       // mini-table editor modal open
  });
});

// ── Init ─────────────────────────────────────────────────────────────────
(async function init() {
  await Promise.all([loadSheet(), loadConfig(), loadStats()]);
  loadKeyboardShortcuts();
  // Check for preload warnings and badge the Logs tab if any exist.
  _checkPreloadWarnings();
})();

async function _checkPreloadWarnings() {
  try {
    const r    = await fetch('/api/logs');
    const data = await r.json();
    const warns = data.preload_warnings || [];
    if (warns.length) {
      const hasErrors = warns.some(w => w.startsWith('[ERROR]'));
      const btn = document.getElementById('tab-logs-btn');
      if (btn) { btn.textContent = hasErrors ? '🔴 Logs' : '⚠️ Logs'; btn.dataset.warned = '1'; }
      _updatePatternErrorBanner(warns);
    }
  } catch (_) { /* ignore */ }
}

function _updatePatternErrorBanner(warns) {
  const errors = warns.filter(w => w.startsWith('[ERROR]'));
  let banner = document.getElementById('pattern-error-banner');
  if (!errors.length) {
    if (banner) banner.remove();
    return;
  }
  if (!banner) {
    banner = document.createElement('div');
    banner.id = 'pattern-error-banner';
    banner.style.cssText = 'background:#5c1a1a;color:#ffcccc;font-size:12px;padding:6px 12px;border-bottom:1px solid #a33;display:flex;gap:8px;align-items:flex-start;';
    const toolbar = document.querySelector('.toolbar');
    if (toolbar) toolbar.after(banner);
  }
  banner.innerHTML = '🔴 <b>Pattern errors</b> — cannot extract until fixed: ' +
    errors.map(e => `<span style="opacity:.85">${_esc(e.replace('[ERROR] ',''))}</span>`).join(' · ');
}

// ── Data fetching ─────────────────────────────────────────────────────────
async function loadSheet() {
  const r = await fetch('/api/sheet');
  sheetData = await r.json();
  renderGrid(sheetData);
  // Index cell data
  for (const row of sheetData.rows)
    for (const cell of row)
      cellData[cell.ref] = cell;
}

async function loadConfig() {
  const r = await fetch('/api/state');
  const s = await r.json();
  populateConfigPanel(s.config);
}

async function loadStats() {
  const r = await fetch('/api/state');
  const s = await r.json();
  currentStats = s.stats;
  renderStats(s.stats);
}

// ── Table role badge builder ──────────────────────────────────────────────
/**
 * Uniform Unicode glyph badge system.  Every T-family cell shows:
 *   [▶]  ▦  [▀|▬|▄|➥]  [L|V|⬚]  [◀]
 *
 * Full matrix (choice × row_class × table_role):
 *   T      header  L/V/I  →  ▶ ▦ ▀ L|V|⬚
 *   T-HEAD header  L/V/I  →    ▦ ▀ L|V|⬚
 *   T-HEAD footer  L/V/I  →    ▦ ▄ L|V|⬚  [◀ if is_table_end]
 *   T-DATA data    L/V/I  →    ▦ ▬ L|V|⬚  [◀ if is_table_end]
 *   L/V/C/I                →  plain choice badge (unchanged)
 *
 * Glyphs:  ▶=table-start  ▦=in-table  ▀=header  ▬=data  ▄=footer
 *          ➥=skip  L=label  V=variable  ⬚=ignored  ◀=table-end
 */
function _tableBadge(choice, table_role, row_class, is_table_end) {
  if (!choice) return '';
  const B  = (cls, txt) => `<span class="cell-badge ${cls}">${txt}</span>`;
  const Ph = txt => `<span class="cell-badge badge-phantom">${txt}</span>`;  // invisible placeholder
  const roleBadge = r => r === 'L' ? B('badge-role-L', 'L')
                       : r === 'V' ? B('badge-role-V', 'V')
                       :             B('badge-role-empty', '⬚');
  // Non-table cells: just the single choice badge (L/V/C/I/T)
  if (!['T', 'T-HEAD', 'T-DATA'].includes(choice)) return B(`badge-${choice}`, choice);
  // Table cells: always 4 slots so text aligns across the column
  //   slot 1: ▶ (anchor only) or phantom for non-anchor
  //   slot 2: ▦ (always)
  //   slot 3: ▀/▬/▄/➥ (row class)
  //   slot 4: L/V/⬚ (role)
  //   [optional] ◀ on the bottom-right (table-end) cell
  const anchor  = choice === 'T' ? B('badge-tbl-anchor', '▶') : Ph('▶');
  const member  = B('badge-tbl-member', '▦');
  const rowG    = row_class === 'header' ? B('badge-row-H', '▀')
                : row_class === 'data'   ? B('badge-row-D', '▬')
                : row_class === 'footer' ? B('badge-row-F', '▄')
                : row_class === 'skip'   ? B('badge-row-S', '➥')
                :                          Ph('▀');
  const role    = roleBadge(table_role || 'I');
  const endMark = is_table_end ? B('badge-tbl-anchor', '◀') : '';
  return anchor + member + rowG + role + endMark;
}

// ── AG Grid cell renderer ─────────────────────────────────────────────────
class WizardCellRenderer {
  init(params) {
    this.eGui = document.createElement('div');
    this.eGui.style.cssText = 'width:100%;height:100%;overflow:hidden;display:flex;align-items:center';
    this._update(params.value);
  }
  _update(cell) {
    if (!cell) { this.eGui.innerHTML = ''; return; }
    if (cell.skip) {
      this.eGui.innerHTML = '<span class="gx-merge-indicator" title="Part of a merged cell area">⊞</span>';
      return;
    }
    const note = cell.note ? ' 📝' : '';
    this.eGui.innerHTML = _cellContent(
      cell.choice, cell.table_role || '', cell.row_class || '',
      cell.is_table_end || false, cell.value || '', note, cell.type || 'str'
    );
  }
  getGui() { return this.eGui; }
  refresh(params) { this._update(params.value); return true; }
  destroy() {}
}

// ── Cell content builder ──────────────────────────────────────────────────
// Returns the innerHTML for a grid cell: badge-strip + cell-value wrapper.
// badge-strip has a fixed width for table cells so the text starts at the
// same x position in every cell of the same column.
function _cellContent(choice, table_role, row_class, is_table_end, valRaw, note, cellType) {
  const badge = _tableBadge(choice, table_role, row_class, is_table_end);
  const val   = escHtml(valRaw);
  const isRight = (cellType === 'num' || cellType === 'number' ||
                   cellType === 'date' ||
                   cellType === 'bool' || cellType === 'boolean');
  const alignCls = isRight ? ' v-right' : '';
  if (!badge) {
    return `<div class="gx-inner"><span class="cell-value${alignCls}">${val}${note}</span></div>`;
  }
  const isTbl = ['T','T-HEAD','T-DATA'].includes(choice);
  const stripCls = isTbl ? 'badge-strip tbl' : 'badge-strip role';
  return `<div class="gx-inner"><span class="${stripCls}">${badge}</span><span class="cell-value${alignCls}">${val}${note}</span></div>`;
}

// ── AG Grid helpers ────────────────────────────────────────────────────────
function _parseColRow(ref) {
  const m = ref.match(/^([A-Z]+)(\d+)$/);
  return m ? [m[1], m[2]] : [null, null];
}

function _refreshCellsByRef(refs) {
  if (!gridApi || !refs.length) return;
  const byRow = {};
  for (const ref of refs) {
    if (!ref) continue;
    const [col, rowStr] = _parseColRow(ref);
    if (!col) continue;
    (byRow[rowStr] = byRow[rowStr] || []).push(col);
  }
  for (const [rowStr, cols] of Object.entries(byRow)) {
    const node = gridApi.getRowNode(rowStr);
    if (node) gridApi.refreshCells({ rowNodes: [node], columns: cols, force: true });
  }
}

// ── Grid rendering ────────────────────────────────────────────────────────
function renderGrid(data) {
  if (!data?.rows) return;

  const colDefs = [
    {
      headerName: '', field: '_rowNum',
      pinned: 'left', width: 40, minWidth: 40, maxWidth: 40,
      suppressMovable: true, resizable: false, sortable: false,
      cellClass: 'gx-row-num-cell',
      cellRenderer: p => String(p.value),
    },
    ...data.col_letters.map(ltr => ({
      headerName: ltr, field: ltr,
      width: 120, minWidth: 40,
      suppressMovable: true, sortable: false,
      cellRenderer: WizardCellRenderer,
      tooltipValueGetter: p => {
        const c = p.value;
        if (!c) return '';
        if (c.skip) return 'Part of a merged cell area';
        return c.name ? `${c.ref}: ${c.value} · field: ${c.name}` : `${c.ref}: ${c.value}`;
      },
      cellClassRules: {
        'gx-cell':             () => true,
        'gx-L':                p => p.value?.choice === 'L',
        'gx-V':                p => p.value?.choice === 'V',
        'gx-T':                p => ['T','T-HEAD','T-DATA'].includes(p.value?.choice),
        'gx-T-HEAD':           p => p.value?.choice === 'T-HEAD',
        'gx-T-DATA':           p => p.value?.choice === 'T-DATA',
        'gx-I':                p => p.value?.choice === 'I',
        'gx-merged-ghost':     p => p.value?.skip === true,
        'gx-image':            p => !p.value?.skip && !!p.value?.has_image && !p.value?.image_suspicious,
        'gx-image-suspicious': p => !p.value?.skip && !!p.value?.image_suspicious,
        'empty':               p => !p.value?.skip && !p.value?.choice && !!p.value?.empty,
        'selected':            p => !!p.value?.ref && p.value.ref === selectedRef,
        'in-range':            p => !!p.value?.ref && _inRangeRefs.has(p.value.ref),
      },
    })),
  ];

  const rowData = data.rows.map(row => {
    const obj = { _rowNum: row[0]?.row || 0 };
    for (const cell of row) {
      if (cell.ref) {
        const col = cell.ref.match(/^[A-Z]+/)?.[0];
        if (col) obj[col] = cell;
      }
    }
    return obj;
  });

  if (gridApi) {
    gridApi.setGridOption('columnDefs', colDefs);
    gridApi.setGridOption('rowData', rowData);
    return;
  }

  const container = document.getElementById('ag-grid-container');
  gridApi = agGrid.createGrid(container, {
    theme: agGrid.themeBalham,
    columnDefs: colDefs,
    rowData: rowData,
    rowHeight: 28,
    headerHeight: 26,
    getRowId: p => String(p.data._rowNum),
    defaultColDef: { resizable: true, sortable: false, filter: false, cellDataType: false },
    suppressCellFocus: true,
    suppressMovableColumns: true,
    tooltipShowDelay: 600,
    onCellClicked: params => {
      const cell = params.value;
      if (!cell || typeof cell !== 'object') return;
      if (cell.skip) { toast(`Cell ${cell.ref} is part of a merged cell area`); return; }
      const ref = cell.ref;
      if (params.event.shiftKey && selectedAction === 'T') {
        handleShiftClick(ref);
      } else if (params.event.shiftKey && selectedRef) {
        _rangeAnchor = _rangeAnchor || selectedRef;
        _rangeEnd = ref;
        _renderRangeHighlight();
      } else {
        selectCell(ref);
      }
    },
  });
}

function refreshCell(ref) {
  if (!gridApi) return;
  const [col, rowStr] = _parseColRow(ref);
  if (!col) return;
  const node = gridApi.getRowNode(rowStr);
  if (!node) return;
  const info = cellData[ref] || {};
  const oldCell = node.data[col] || {};
  node.setDataValue(col, { ...oldCell, ...info });
}

// ── Cell selection ─────────────────────────────────────────────────────────
async function selectCell(ref) {
  // Clear any active range selection and T-mode anchor when clicking normally
  _clearRange();
  _tModeAnchor = null;
  const oldRef = selectedRef;
  selectedRef = ref;
  _refreshCellsByRef([oldRef, ref].filter(Boolean));

  // Fetch full cell info
  const r = await fetch(`/api/cell/${ref}`);
  const info = await r.json();
  cellData[ref] = { ...cellData[ref], ...info };

  // Update panel header
  document.getElementById('ph-ref').textContent = ref;
  document.getElementById('ph-val').textContent =
    info.raw ? `"${info.raw.substring(0,80)}"  [${info.inferred_type}]` : '(empty cell)';

  // Show classify form
  Alpine.store('gx').panelMode = 'classify';

  // Pre-fill form from existing classification
  prefillForm(info);
  _updateCellDetail(ref, info);
  setStatus(`Selected ${ref}` + (info.choice ? ` — currently ${info.choice}` : ''));
}

function _updateCellDetail(ref, info) {
  const box = document.getElementById('cell-detail-content');
  if (!box) return;
  const TYPE_COLOR = {
    string:'#6ea8d0', integer:'#4ade80', number:'#4ade80', date:'#fb923c',
    boolean:'#c084fc', url:'#38bdf8', image:'#2dd4bf', empty:'var(--text-muted)',
  };
  const rows = [];

  // Raw value (full, scrollable)
  if (info.raw) {
    rows.push(`<div style="margin-bottom:7px">
      <span style="color:var(--text-muted);font-size:11px">Value</span>
      <div style="margin-top:2px;font-family:monospace;font-size:11px;word-break:break-all;max-height:60px;overflow-y:auto;background:var(--surface);padding:3px 5px;border-radius:3px">${_esc(String(info.raw).substring(0,400))}</div>
    </div>`);
  }

  // Inferred type badge
  const itype = info.inferred_type || 'empty';
  const tc = TYPE_COLOR[itype] || '#888';
  rows.push(`<div style="margin-bottom:7px">
    <span style="color:var(--text-muted);font-size:11px">Type</span>
    <span style="margin-left:6px;padding:1px 7px;border-radius:10px;font-size:11px;font-weight:600;background:${tc}22;color:${tc}">${_esc(itype)}</span>
  </div>`);

  // Image info + preview
  if (info.has_image) {
    const susp = info.image_suspicious;
    const mimes = (info.image_mimes || []).filter(m => m !== 'image/embedded').join(', ') || 'embedded';
    const preview = info.has_preview
      ? `<img src="/api/image/${_esc(ref)}" alt="cell image"
              style="max-width:100%;max-height:120px;margin-top:6px;border-radius:4px;border:1px solid var(--border);display:block"
              onerror="this.style.display='none'">`
      : (susp ? '' : '<div style="font-size:11px;color:var(--text-muted);margin-top:4px">Formula image — no binary to preview</div>');
    rows.push(`<div style="margin-bottom:7px;padding:6px 8px;border-radius:5px;background:${susp?'rgba(245,158,11,.12)':'rgba(20,184,166,.12)'};border:1px solid ${susp?'rgba(245,158,11,.4)':'rgba(20,184,166,.3)'}">
      <div style="font-weight:600;color:${susp?'#f59e0b':'#2dd4bf'}">${susp?'⚠️':'🖼'} ${info.image_count} image${info.image_count!==1?'s':''}</div>
      <div style="font-size:11px;color:var(--text-muted);margin-top:2px">${_esc(mimes)}</div>
      ${susp?'<div style="font-size:11px;color:#f59e0b;margin-top:3px">Suspicious content — not extracted to disk</div>':''}
      ${preview}
    </div>`);
  }

  // Current classification
  if (info.choice) {
    rows.push(`<div style="margin-bottom:7px">
      <span style="color:var(--text-muted);font-size:11px">Classified</span>
      <span style="margin-left:6px;font-size:11px;font-weight:600">${_esc(info.choice)}</span>
      ${info.name ? `<code style="margin-left:4px;font-size:11px;padding:0 4px;background:var(--surface);border-radius:3px">${_esc(info.name)}</code>` : ''}
    </div>`);
  }

  // Position
  rows.push(`<div style="color:var(--text-muted);font-size:11px">${_esc(ref)} · Row ${info.row} · Col ${info.col}</div>`);

  box.innerHTML = rows.length ? rows.join('') : '<em style="color:var(--text-muted)">Empty cell</em>';
}

function _enterTableMode(anchorRef, clickedRef) {
  Alpine.store('gx').panelMode      = 'table';
  Alpine.store('gx').selectedAction = 'T';

  // Move selection highlight to the anchor
  const prevRef = selectedRef;
  selectedAction = 'T';
  selectedRef    = anchorRef;
  _refreshCellsByRef([prevRef, anchorRef, clickedRef].filter(Boolean));
}

function _exitTableMode() {
  Alpine.store('gx').panelMode = 'classify';
}

function prefillForm(info) {
  const ch = info.choice || '';

  // Table cells get a restricted panel — no classify buttons
  if (ch === 'T' || ch === 'T-HEAD' || ch === 'T-DATA') {
    const anchorRef = (ch === 'T') ? (info.ref || selectedRef) : (info.anchor || '');
    if (!anchorRef) { _exitTableMode(); return; }

    _enterTableMode(anchorRef, info.ref || selectedRef);

    document.getElementById('ph-val').textContent =
      ch === 'T' ? `Mini Table anchor` : `[${ch}] part of table at ${anchorRef}`;

    return;
  }

  // Normal (non-table) cell — restore classify buttons
  _exitTableMode();

  // Reset all action buttons
  document.querySelectorAll('.cbtn').forEach(b => b.classList.remove('active'));
  if (ch && ch !== 'CLEAR') {
    const btn = document.querySelector(`.cbtn-${ch}`);
    if (btn) btn.classList.add('active');
  }

  if (!ch) {
    selectedAction = null;
    Alpine.store('gx').selectedAction = null;
    return;
  }
  selectedAction = ch;
  showFields(ch);

  if (ch === 'L') {
    setVal('f-L-name', info.name || '');
    setVal('f-L-type', info.ltype || 'string');
    setVal('f-L-match_mode', info.lbl_mode ? info.lbl_mode : '(default)');
    setVal('f-L-match', info.lmatch || '');
    setVal('f-L-notes', info.note || '');
  } else if (ch === 'V') {
    setVal('f-V-name', info.name || '');
    setVal('f-V-type', info.ftype || 'string');
    setVal('f-V-match_mode', info.var_mode || '(default)');
    setVal('f-V-match', info.match || '.*');
    setVal('f-V-modifiers', info.modifiers || 'none');
    setVal('f-V-notes', info.note || '');
    onVTypeChange(info.ftype || 'string');
  }
}

async function _clearRefs(anchorRef) {
  // Send CLEAR to the backend; response includes all cascaded refs
  const data = await classify(anchorRef, 'CLEAR', {});
  const cleared = (data && data.cleared) ? data.cleared : [anchorRef];
  // Wipe cellData cache and refresh grid for every cleared cell
  for (const r of cleared) {
    const cached = cellData[r] || {};
    cached.choice = ''; cached.name = ''; cached.anchor = '';
    cellData[r] = cached;
    refreshCell(r);
  }
  await loadStats();
  return cleared;
}

async function removeTable() {
  if (!selectedRef) return;
  await _clearRefs(selectedRef);
  _exitTableMode();
  prefillForm({});
  toast('Mini-Table removed');
}

function selectAction(action) {
  selectedAction = action;
  document.querySelectorAll('.cbtn').forEach(b => b.classList.remove('active'));
  const btn = document.querySelector(`.cbtn-${action}`);
  if (btn) btn.classList.add('active');
  document.querySelectorAll('.action-fields').forEach(d => d.style.display = 'none');
  // Clear T-mode anchor when switching away from T
  if (action !== 'T') _tModeAnchor = null;

  if (!selectedRef) return;
  const info = cellData[selectedRef] || {};

  // Default pre-fill
  if (action === 'L') {
    const slug = slugify(info.raw || 'field');
    setVal('f-L-name', info.name || (slug + '_label'));
    setVal('f-L-type', info.ltype || 'string');
    setVal('f-L-match_mode', '(default)');
    setVal('f-L-match', info.raw || '');
    setVal('f-L-notes', info.note || '');
  } else if (action === 'V') {
    const suggested = suggestNameFromLabel(selectedRef);
    const inferredType = info.inferred_type || info.ftype || 'string';
    setVal('f-V-name', info.name || suggested || slugify(info.raw || 'field'));
    setVal('f-V-type', inferredType);
    setVal('f-V-match_mode', '(default)');
    setVal('f-V-match', info.match || '.*');
    setVal('f-V-modifiers', 'none');
    setVal('f-V-notes', info.note || '');
    onVTypeChange(inferredType);
  } else if (action === 'T') {
    const defaultName = info.name || (slugify(info.raw || 'table') + '_table');
    setVal('f-T-name', defaultName);
    setVal('f-T-mult', info.mult || '*');
    // If this cell already has table context, load end_ref + row_config status
    if (info.choice === 'T') {
      fetch(`/api/table-context/${selectedRef}`)
        .then(r => r.json())
        .then(ctx => {
          if (ctx.end_ref && ctx.end_ref !== selectedRef) {
            setVal('f-T-end_ref', ctx.end_ref);
          } else {
            const inferred = _inferTableEndRef(selectedRef);
            if (inferred) {
              setVal('f-T-end_ref', inferred);
            } else if (_rangeAnchor === selectedRef && _rangeEnd) {
              // Active range selection covers this anchor — use it as end_ref
              setVal('f-T-end_ref', _rangeEnd.toUpperCase());
              _tModeAnchor = selectedRef;
            }
          }
          _tUpdateRowConfigStatus(ctx);
        })
        .catch(() => {});
    } else {
      // For a fresh anchor: pre-populate end_ref from any active range selection
      // so the user can: select range with Shift+Arrow → press T → Enter opens editor.
      const rangeEndForT = (_rangeAnchor === selectedRef && _rangeEnd) ? _rangeEnd : '';
      setVal('f-T-end_ref', rangeEndForT.toUpperCase());
      if (rangeEndForT) {
        // Range is now "claimed" by T mode — keep the visual highlight but
        // remember the anchor so openTableModal uses it correctly.
        _tModeAnchor = selectedRef;
      }
      const s = document.getElementById('t-row-config-status');
      if (s) s.style.display = 'none';
    }
    // Auto-scan columns from the anchor row (legacy column list)
    _tScanColumns(selectedRef, []);
  }

  showFields(action);
}

function showFields(action) {
  Alpine.store('gx').selectedAction = action;
}

function cancelClassification() {
  selectedAction = null;
  document.querySelectorAll('.cbtn').forEach(b => b.classList.remove('active'));
  Alpine.store('gx').selectedAction = null;
  Alpine.store('gx').panelMode      = selectedRef ? 'classify' : 'hint';
}

async function clearCell() {
  if (!selectedRef) return;
  const cleared = await _clearRefs(selectedRef);
  prefillForm({});
  toast(`Cleared ${cleared.length > 1 ? cleared.length + ' cells' : selectedRef}`);
}

async function applyClassification() {
  if (!selectedRef || !selectedAction) return;

  // T action must go through the Row Editor modal — never the legacy sidebar Apply path
  if (selectedAction === 'T') {
    openTableModal();
    return;
  }

  const fields = gatherFields(selectedAction);
  const classifyData = await classify(selectedRef, selectedAction, fields);
  const info = cellData[selectedRef] || {};
  info.choice = selectedAction;
  info.name   = fields.name || '';
  info.note   = fields.notes || '';

  refreshCell(selectedRef);
  toast(`Classified ${selectedRef} as ${selectedAction}`);
  if (classifyData?.name_warning) {
    setTimeout(() => toast('⚠️ ' + classifyData.name_warning, false, 8000), 400);
  }
  setStatus(`${selectedRef} → ${selectedAction} (${fields.name || ''})`);
  await loadStats();
}

function gatherFields(action) {
  const g = id => (document.getElementById(id) || {}).value || '';
  if (action === 'L') return { name: g('f-L-name'), type: g('f-L-type'), match_mode: g('f-L-match_mode'), match: g('f-L-match'), notes: g('f-L-notes') };
  if (action === 'V') return { name: g('f-V-name'), type: g('f-V-type'), match_mode: g('f-V-match_mode'), match: g('f-V-match'), modifiers: g('f-V-modifiers'), notes: g('f-V-notes') };
  if (action === 'T') {
    const mult   = g('f-T-mult') || '*';
    const endRef = (g('f-T-end_ref') || '').toUpperCase().trim();
    // If the Row Editor modal was used, row_configs are stored on the anchor;
    // the backend classify handler picks them up. For legacy column-list flow:
    const colRows = document.querySelectorAll('#t-col-list .t-col-row');
    const columns = Array.from(colRows).map(row => ({
      ref:        row.dataset.ref  || '',
      header:     row.dataset.header || '',
      field_name: row.querySelector('.t-col-field')?.value?.trim() || 'IGNORE',
      var_type:   'string',
      var_match:  '.*',
    }));
    return { name: g('f-T-name'), mult, end_ref: endRef, columns };
  }
  if (action === 'I') return {};
  return {};
}

async function classify(ref, action, fields) {
  const r = await fetch('/api/classify', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ ref, action, fields }),
  });
  const data = await r.json();
  if (data.stats) { currentStats = data.stats; renderStats(currentStats); }
  return data;
}

async function undoLast() {
  const r = await fetch('/api/undo', { method: 'POST' });
  const data = await r.json();
  if (data.ok) {
    toast('Undid ' + (data.ref || 'last action'));
    if (data.stats) { currentStats = data.stats; renderStats(currentStats); }
    await loadSheet();
    if (selectedRef) {
      const re2 = await fetch(`/api/cell/${selectedRef}`);
      const info = await re2.json();
      cellData[selectedRef] = { ...cellData[selectedRef], ...info };
      prefillForm(info);
      refreshCell(selectedRef);
    }
  } else {
    toast(data.msg || 'Nothing to undo');
  }
}

// ── Stats bar (always-visible bottom strip) ───────────────────────────────
function renderStats(stats) {
  // All typed cells count as classified: L+V+C+T+T-HEAD+T-DATA+I
  const tHead = stats.T_HEAD || 0;
  const tData = stats.T_DATA || 0;
  const classified   = (stats.L||0)+(stats.V||0)+(stats.C||0)+(stats.T||0)+tHead+tData+(stats.I||0);
  const unclassified = stats.unclassified != null ? stats.unclassified : Math.max(0, (stats.total||0) - classified);
  const total        = stats.total || 0;
  const pct          = total > 0 ? Math.round((classified / total) * 100) : 0;

  const setText = (id, v) => { const el = document.getElementById(id); if (el) el.textContent = v; };
  setText('sb-L', stats.L || 0);
  setText('sb-V', stats.V || 0);
  setText('sb-T', stats.T || 0);
  setText('sb-I', stats.I || 0);
  setText('sb-U', unclassified);
  setText('sb-pct', `${classified} / ${total} classified (${pct}%)`);

  // Show T-HEAD / T-DATA sub-counts inside the Mini Tables chip
  const detail = document.getElementById('sb-T-detail');
  if (detail) {
    const parts = [];
    if (tHead > 0) parts.push(`${tHead} col${tHead !== 1 ? 's' : ''}`);
    if (tData > 0) parts.push(`${tData} rows`);
    detail.textContent = parts.length ? ' · ' + parts.join(' · ') : '';
  }

  const fill = document.getElementById('sb-fill');
  if (fill) fill.style.width = Math.min(pct, 100) + '%';

  // Highlight Unclassified chip in amber when there are pending cells
  const uChip = document.getElementById('sb-U-chip');
  if (uChip) uChip.classList.toggle('has-pending', unclassified > 0);
}

// ── Config ────────────────────────────────────────────────────────────────
function _cfgSelMark(sel, defaultVal) {
  sel.classList.toggle('cfg-sel-default', sel.value === defaultVal);
}

function populateConfigPanel(cfg) {
  setVal('cfg-direction',         cfg.direction  || 'LR');
  setChk('cfg-template',          cfg.template   || false);
  // split ignore_case flags (server returns split keys; fall back to legacy ignore_case for old sessions)
  const legacyIC = cfg.ignore_case || false;
  setChk('cfg-ignore_case_labels', cfg.ignore_case_labels != null ? cfg.ignore_case_labels : legacyIC);
  setChk('cfg-ignore_case_values', cfg.ignore_case_values != null ? cfg.ignore_case_values : legacyIC);
  // split trim_whitespace flags
  const legacyTW = cfg.trim_whitespace || false;
  setChk('cfg-trim_ws_labels',    cfg.trim_whitespace_labels != null ? cfg.trim_whitespace_labels : legacyTW);
  setChk('cfg-trim_ws_values',    cfg.trim_whitespace_values != null ? cfg.trim_whitespace_values : legacyTW);
  setVal('cfg-currency',          cfg.currency_sign || '€');
  setVal('cfg-lbl_match',         cfg.lbl_match  || '');
  setVal('cfg-var_match',         cfg.var_match  || '');
  setVal('cfg-aliases',           (cfg.empty_aliases || []).join(', '));
  _cfgSelMark(document.getElementById('cfg-direction'), 'LR');
  _cfgSelMark(document.getElementById('cfg-lbl_match'), '');
  _cfgSelMark(document.getElementById('cfg-var_match'), '');
}

async function saveConfig() {
  const aliases = document.getElementById('cfg-aliases').value
    .split(',').map(s => s.trim()).filter(Boolean);
  const cfg = {
    direction:             document.getElementById('cfg-direction').value,
    template:              document.getElementById('cfg-template').checked,
    ignore_case_labels:    document.getElementById('cfg-ignore_case_labels').checked,
    ignore_case_values:    document.getElementById('cfg-ignore_case_values').checked,
    trim_whitespace_labels: document.getElementById('cfg-trim_ws_labels').checked,
    trim_whitespace_values: document.getElementById('cfg-trim_ws_values').checked,
    currency_sign:         document.getElementById('cfg-currency').value || '€',
    lbl_match:             document.getElementById('cfg-lbl_match').value,
    var_match:             document.getElementById('cfg-var_match').value,
    empty_aliases:         aliases,
  };
  await fetch('/api/config', { method: 'POST', headers: {'Content-Type':'application/json'}, body: JSON.stringify(cfg) });
  toast('Config saved');
  setStatus('Config updated');
}

// ── Preview & Save ─────────────────────────────────────────────────────────
async function showPreview() {
  const modal = document.getElementById('preview-modal');
  const pre   = document.getElementById('preview-text');
  pre.textContent = 'Loading…';   // reset each time so stale content never shows
  modal.classList.add('open');
  try {
    const r = await fetch('/api/preview');
    if (!r.ok) { pre.textContent = `Error ${r.status}: ${await r.text()}`; return; }
    csvText = await r.text();
    pre.textContent = csvText;
  } catch (e) {
    pre.textContent = `Network error: ${e.message}`;
  }
}
function closePreview() { document.getElementById('preview-modal').classList.remove('open'); }
function downloadCSV() {
  const blob = new Blob([csvText], {type: 'text/csv'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'pattern.csv';
  a.click();
}

async function uploadPattern(input) {
  const file = input.files[0];
  if (!file) return;
  input.value = '';  // reset so the same file can be re-selected
  const fd = new FormData();
  fd.append('file', file);
  toast(`Loading pattern from ${file.name}…`);
  try {
    const r = await fetch('/api/load-pattern', { method: 'POST', body: fd });
    const data = await r.json();
    if (data.ok) {
      const warn = data.warnings && data.warnings.length
        ? ` (${data.warnings.length} warning(s))` : '';
      toast(`Pattern loaded — ${data.n_choices} cells classified${warn}. Reloading grid…`);
      // Short delay so the user can read the toast, then reload
      setTimeout(() => window.location.reload(), 900);
    } else {
      toast('Load failed: ' + (data.detail || data.error || '?'), true);
    }
  } catch (err) {
    toast('Load failed: ' + err, true);
  }
}

async function savePattern(confirmOverwrite) {
  const body = confirmOverwrite ? {confirm_overwrite: true} : {};
  const r = await fetch('/api/save', {
    method: 'POST',
    headers: {'Content-Type':'application/json'},
    body: JSON.stringify(body),
  });
  const data = await r.json();
  if (data.ok) {
    toast(`Saved CSV → ${data.saved_to}`);
    setStatus(`Pattern saved → ${data.saved_to}`);
  } else if (data.exists) {
    if (confirm(`File already exists:\n${data.path}\n\nOverwrite?`)) savePattern(true);
  } else {
    toast('Save failed: ' + (data.detail || '?'), true);
  }
}

async function savePatternXlsx(confirmOverwrite) {
  const body = confirmOverwrite ? {confirm_overwrite: true} : {};
  const r = await fetch('/api/save-xlsx', {
    method: 'POST',
    headers: {'Content-Type':'application/json'},
    body: JSON.stringify(body),
  });
  const data = await r.json();
  if (data.ok) {
    toast(`Saved Excel → ${data.saved_to}`);
    setStatus(`Pattern saved → ${data.saved_to}`);
  } else if (data.exists) {
    if (confirm(`File already exists:\n${data.path}\n\nOverwrite?`)) savePatternXlsx(true);
  } else {
    toast('Save Excel failed: ' + (data.detail || '?'), true);
  }
}

// ── Sheet switching ────────────────────────────────────────────────────────
let _sheetList = [];
async function toggleSheetSwitcher() {
  const menu = document.getElementById('sheet-switcher-menu');
  if (menu.style.display !== 'none') { menu.style.display = 'none'; return; }
  // Close other menus
  document.getElementById('load-pattern-menu').style.display = 'none';
  try {
    const r = await fetch('/api/sheets');
    const d = await r.json();
    _sheetList = d.sheets || [];
    const active = d.active || '';
    const list = document.getElementById('sheet-switcher-list');
    list.innerHTML = _sheetList.map(s => {
      const isCur = s === active;
      return `<div style="padding:5px 12px;cursor:${isCur?'default':'pointer'};background:${isCur?'var(--accent-dim)':'transparent'};font-weight:${isCur?'600':'400'}"
                   data-sheet="${_esc(s)}" ${isCur ? '' : 'onclick="doSwitchSheet(this.dataset.sheet)"'}>
                ${_esc(s)}${isCur?' ✓':''}
              </div>`;
    }).join('');
    const btnRect = document.getElementById('sheet-switcher-btn').getBoundingClientRect();
    menu.style.top   = (btnRect.bottom + 4) + 'px';
    menu.style.right = (window.innerWidth - btnRect.right) + 'px';
    menu.style.display = 'block';
  } catch(e) { toast('Could not load sheets: ' + e, true); }
}

async function doSwitchSheet(name) {
  document.getElementById('sheet-switcher-menu').style.display = 'none';
  const r = await fetch('/api/switch-sheet', {
    method: 'POST',
    headers: {'Content-Type':'application/json'},
    body: JSON.stringify({sheet: name}),
  });
  const d = await r.json();
  if (!d.ok) { toast('Switch failed: ' + (d.detail || '?'), true); return; }
  // Update navbar sheet label
  const el = document.getElementById('navbar-sheet');
  if (el) el.textContent = name;
  // Reset local state and reload grid
  cellData = {};
  selectedRef = null;
  sheetData = null;
  await loadSheet();
  await loadStats();
  toast(`Switched to sheet: ${name}`);
}

// ── Load Pattern from folder ───────────────────────────────────────────────
async function toggleLoadPatternMenu() {
  const menu = document.getElementById('load-pattern-menu');
  if (menu.style.display !== 'none') { menu.style.display = 'none'; return; }
  // Close other menus
  document.getElementById('sheet-switcher-menu').style.display = 'none';
  // Load folder patterns
  const list = document.getElementById('load-pattern-folder-list');
  list.innerHTML = '<em style="color:var(--text-muted);font-size:12px">loading…</em>';
  const lpBtn = document.getElementById('load-pattern-btn');
  const lpRect = lpBtn.getBoundingClientRect();
  menu.style.top   = (lpRect.bottom + 4) + 'px';
  menu.style.right = (window.innerWidth - lpRect.right) + 'px';
  menu.style.display = 'block';
  try {
    const r = await fetch('/api/list-patterns');
    const d = await r.json();
    const patterns = d.patterns || [];
    if (!patterns.length) {
      list.innerHTML = '<em style="color:var(--text-muted);font-size:12px">No pattern files found</em>';
    } else {
      list.innerHTML = patterns.map(p =>
        `<div style="padding:4px 8px;cursor:pointer;border-radius:3px" class="folder-pat-item"
              data-path="${_esc(p.path)}" data-name="${_esc(p.name)}"
              onclick="loadPatternFromFolder(this.dataset.path,this.dataset.name)"
              onmouseover="this.style.background='var(--accent-dim)'" onmouseout="this.style.background=''">
           <div style="font-size:12px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${_esc(p.name)}</div>
           <div style="font-size:10px;color:var(--text-muted)">${p.size ? Math.round(p.size/1024)+' KB' : '—'}</div>
         </div>`
      ).join('');
    }
  } catch(e) { list.innerHTML = `<em style="color:var(--text-muted);font-size:12px">Error: ${_esc(String(e))}</em>`; }
}

async function loadPatternFromFolder(path, name) {
  document.getElementById('load-pattern-menu').style.display = 'none';
  try {
    const r = await fetch('/api/load-pattern-by-path', {
      method: 'POST',
      headers: {'Content-Type':'application/json'},
      body: JSON.stringify({path}),
    });
    const d = await r.json();
    if (!d.ok) { toast('Load failed: ' + (d.detail || '?'), true); return; }
    cellData = {}; selectedRef = null; sheetData = null;
    await loadSheet(); await loadStats();
    if (d.warnings?.length) toast(`Loaded "${name}" with ${d.warnings.length} warning(s)`);
    else toast(`Loaded pattern: ${name}`);
  } catch(e) { toast('Load failed: ' + e, true); }
}

// Close dropdowns when clicking outside
document.addEventListener('click', e => {
  if (!document.getElementById('sheet-switcher-wrap')?.contains(e.target))
    document.getElementById('sheet-switcher-menu').style.display = 'none';
  if (!document.getElementById('load-pattern-wrap')?.contains(e.target))
    document.getElementById('load-pattern-menu').style.display = 'none';
});

// ── Tabs ──────────────────────────────────────────────────────────────────
function switchTab(name) {
  Alpine.store('gx').activeTab = name;
  if (name === 'logs') refreshLogs();
}

async function refreshLogs() {
  const linesBox = document.getElementById('log-lines-box');
  const warnBox  = document.getElementById('log-warn-box');
  const pathLine = document.getElementById('log-path-line');
  try {
    const r    = await fetch('/api/logs');
    const data = await r.json();

    // ── Preload warnings / errors ─────────────────────────────────────────
    const warns = data.preload_warnings || [];
    const hasErrors = warns.some(w => w.startsWith('[ERROR]'));
    if (warns.length) {
      warnBox.innerHTML = warns.map(w => {
        const isErr = w.startsWith('[ERROR]');
        const label = isErr ? '🔴' : '⚠️';
        const bg    = isErr ? 'var(--err-bg,#5c1a1a)' : 'var(--warn,#7c4a00)';
        return `<div style="background:${bg};color:#ffe;border-radius:4px;padding:5px 8px;margin:0 8px 4px;font-size:11px;">${label} ${_esc(w)}</div>`;
      }).join('');
      warnBox.style.display = 'block';
      // Badge the tab — red for errors, amber for warnings only
      const btn = document.getElementById('tab-logs-btn');
      if (btn && !btn.dataset.warned) {
        btn.dataset.warned = '1';
        btn.textContent = hasErrors ? '🔴 Logs' : '⚠️ Logs';
      }
    } else {
      warnBox.style.display = 'none';
    }
    // Show/update the persistent error banner on the main grid view
    _updatePatternErrorBanner(warns);

    // ── Log lines ─────────────────────────────────────────────────────────
    const lines = data.log_lines || [];
    if (lines.length) {
      linesBox.textContent = lines.join('\n');
      linesBox.scrollTop   = linesBox.scrollHeight;
    } else {
      linesBox.textContent = '(no log entries yet)';
    }

    // ── Log path ──────────────────────────────────────────────────────────
    if (data.log_path) {
      pathLine.textContent = '📁 ' + data.log_path;
    }
  } catch (e) {
    if (linesBox) linesBox.textContent = 'Error loading logs: ' + e.message;
  }
}


// ── Extract panel ─────────────────────────────────────────────────────────
let _extractProvenance = {};

async function runExtraction() {
  const btn    = document.getElementById('ext-run-btn');
  const status = document.getElementById('ext-status');
  const result = document.getElementById('ext-result');
  btn.disabled = true;
  btn.textContent = '⏳ Running…';
  status.textContent = '';
  result.innerHTML = '';
  try {
    const r = await fetch('/api/extract', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
    });
    const data = await r.json();
    if (!data.ok) {
      result.innerHTML = `<div class="et-error">❌ ${escHtml(data.error || 'Extraction failed')}</div>`;
      status.textContent = 'Extraction failed';
      return;
    }
    _extractProvenance = data.provenance || {};
    result.innerHTML = renderExtractTree(data.result);
    // Attach range-jump handlers to source refs (set via data-range attr)
    result.querySelectorAll('[data-range]').forEach(el => {
      el.addEventListener('click', () => jumpToCell(el.dataset.range.split(':')[0]));
    });
    // Count extracted scalars/tables for status line
    const scalars = Object.entries(data.result || {})
      .filter(([k, v]) => v !== null && typeof v !== 'object').length;
    const tables = Object.entries(data.result || {})
      .filter(([k, v]) => Array.isArray(v)).length;
    status.textContent = `✓ ${scalars} scalars · ${tables} table group${tables !== 1 ? 's' : ''}`;
  } catch (e) {
    result.innerHTML = `<div class="et-error">❌ ${escHtml(e.message)}</div>`;
    status.textContent = 'Error';
  } finally {
    btn.disabled = false;
    btn.textContent = '▶ Run Extraction';
  }
}

function renderExtractTree(result) {
  if (!result || typeof result !== 'object') return '<span class="et-null">null</span>';
  return renderExtObject(result, '');
}

function renderExtObject(obj, parentPath) {
  const entries = Object.entries(obj);
  if (entries.length === 0) return '<span class="et-null">{}</span>';
  return entries.map(([k, v]) => renderExtEntry(k, v, parentPath ? parentPath + '.' + k : k)).join('');
}

function renderExtEntry(key, val, path) {
  if (key === '_source') return '';          // skip internal source metadata
  if (val === null || val === undefined) {
    return `<div class="et-row">
      <span class="et-key">${escHtml(key)}</span>
      <span class="et-colon">:</span>
      <span class="et-null">null</span>
    </div>`;
  }
  if (Array.isArray(val)) {
    // Table group (array of instances)
    return renderExtTableGroup(key, val, path);
  }
  if (typeof val === 'object') {
    // Nested block
    return `<div class="et-group">
      <div class="et-group-title">📂 ${escHtml(key)}</div>
      <div class="et-group-body">${renderExtObject(val, path)}</div>
    </div>`;
  }
  // Scalar leaf — look up provenance refs
  const leafKey = path.split('.').pop();
  const refs = (_extractProvenance[leafKey] || []).slice(0, 4);
  const refChips = refs.map(r =>
    `<span class="et-ref" onclick="jumpToCell('${r}')">${r}</span>`
  ).join(' ');
  return `<div class="et-row">
    <span class="et-key">${escHtml(key)}</span>
    <span class="et-colon">:</span>
    <span class="et-val">${escHtml(String(val))}</span>
    ${refChips}
  </div>`;
}

function renderExtTableGroup(key, instances, path) {
  if (instances.length === 0) {
    return `<div class="et-group">
      <div class="et-group-title">📊 ${escHtml(key)}</div>
      <div class="et-group-body"><span class="et-null">0 instances</span></div>
    </div>`;
  }
  const items = instances.map((inst, i) => {
    const src = inst._source;
    const srcRef = src ? `<span class="et-source-ref" data-range="${escHtml(src.ref)}"
      title="Jump to ${src.ref}">${escHtml(src.ref)}</span>` : '';
    // Gather data rows
    const dataRows = inst.data || [];
    const headerObj = inst.header || {};
    const footerObj = inst.footer || {};
    // Build header labels from first data row keys
    const colKeys = dataRows.length ? Object.keys(dataRows[0]) : Object.keys(headerObj);
    const headerHtml = Object.keys(headerObj).length
      ? `<div class="et-row" style="font-weight:600;font-size:11px;color:var(--text-muted)">
          header: ${Object.entries(headerObj).map(([k,v]) => `${escHtml(k)}=${escHtml(String(v))}`).join(', ')}
         </div>` : '';
    const rowsHtml = dataRows.slice(0, 5).map(row => {
      const cells = colKeys.map(k => `<span class="et-data-cell">${escHtml(String(row[k] ?? ''))}</span>`).join('');
      return `<div class="et-data-row">${cells}</div>`;
    }).join('');
    const moreHtml = dataRows.length > 5
      ? `<div class="et-count">…${dataRows.length - 5} more rows</div>` : '';
    const footerHtml = Object.keys(footerObj).length
      ? `<div class="et-row" style="font-style:italic;font-size:11px;color:var(--text-muted)">
          footer: ${Object.entries(footerObj).map(([k,v]) => `${escHtml(k)}=${escHtml(String(v))}`).join(', ')}
         </div>` : '';
    return `<div class="et-table-instance">
      <div class="et-instance-hdr">
        <span>Instance ${i + 1}</span>
        <span class="et-count">(${dataRows.length} data rows)</span>
        ${srcRef}
      </div>
      <div class="et-instance-body">
        ${headerHtml}
        <div class="et-data-rows">${rowsHtml}${moreHtml}</div>
        ${footerHtml}
      </div>
    </div>`;
  }).join('');
  return `<div class="et-group">
    <div class="et-group-title">📊 ${escHtml(key)} <span class="et-count">(${instances.length} instance${instances.length !== 1 ? 's' : ''})</span></div>
    <div class="et-group-body">${items}</div>
  </div>`;
}

// ── Jump to cell ──────────────────────────────────────────────────────────
function jumpToCell(ref) {
  if (!ref) return;
  ref = ref.toUpperCase().trim();
  if (!gridApi) { toast('Grid not ready', true); return; }
  const [col, rowStr] = _parseColRow(ref);
  if (!col) { toast(`Cell ${ref} not found in visible grid`, true); return; }
  const node = gridApi.getRowNode(rowStr);
  if (!node) { toast(`Cell ${ref} not found in visible grid`, true); return; }
  gridApi.ensureColumnVisible(col);
  gridApi.ensureIndexVisible(node.rowIndex, 'middle');
  selectCell(ref);
  document.getElementById('cell-jump').value = '';
}

// ── Column-letter helpers (for arrow navigation) ───────────────────────────
function colLetterToNum(letters) {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64;
  return n;
}
function numToColLetter(n) {
  let s = '';
  while (n > 0) { n--; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26); }
  return s;
}


// ── Range selection state ─────────────────────────────────────────────────
// _rangeAnchor: the cell where the range began (stays fixed as end moves)
// _rangeEnd:    the current far corner of the selection
let _rangeAnchor = null;
let _rangeEnd    = null;

/** Return all refs in the rectangular bounding box [anchor..end]. */
function _rangeRefs(anchorRef, endRef) {
  const [r1, c1] = _parseRef(anchorRef);
  const [r2, c2] = _parseRef(endRef);
  const minR = Math.min(r1, r2), maxR = Math.max(r1, r2);
  const minC = Math.min(c1, c2), maxC = Math.max(c1, c2);
  const refs = [];
  for (let r = minR; r <= maxR; r++) {
    for (let c = minC; c <= maxC; c++) {
      refs.push(_buildRef(r, c));
    }
  }
  return refs;
}

/** Apply `in-range` class to all cells in bounding box via AG Grid refresh. */
function _renderRangeHighlight() {
  const oldRefs = _inRangeRefs;
  _inRangeRefs = new Set(_rangeAnchor && _rangeEnd ? _rangeRefs(_rangeAnchor, _rangeEnd) : []);
  const toRefresh = new Set([...oldRefs, ..._inRangeRefs]);
  _refreshCellsByRef([...toRefresh]);
  if (!_rangeAnchor || !_rangeEnd) return;
  const [r1, c1] = _parseRef(_rangeAnchor);
  const [r2, c2] = _parseRef(_rangeEnd);
  const rows = Math.abs(r2 - r1) + 1, cols = Math.abs(c2 - c1) + 1;
  setStatus(`Range ${_rangeAnchor}:${_rangeEnd} — ${rows}×${cols} (${_inRangeRefs.size} cells) · press L/V/C/I to classify, Esc to cancel`);
}

/** Cancel any active range selection. */
function _clearRange() {
  const toRefresh = [..._inRangeRefs];
  _rangeAnchor = null;
  _rangeEnd    = null;
  _inRangeRefs = new Set();
  _refreshCellsByRef(toRefresh);
}

/** Apply a batch classification to the current range.  Returns false if no range active. */
async function _classifyRange(action) {
  if (!_rangeAnchor || !_rangeEnd) return false;
  const refs = _rangeRefs(_rangeAnchor, _rangeEnd);
  if (refs.length === 0) return false;
  const label = action === 'CLEAR' ? 'Clear' : action;
  toast(`Classifying ${refs.length} cells as ${label}…`);
  try {
    const r = await fetch('/api/classify-batch', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({ refs, action }),
    });
    const data = await r.json();
    if (!data.ok) {
      toast('Batch classify failed: ' + (data.detail || data.error || '?'), true);
      return true;
    }
    const skipped = data.skipped && data.skipped.length ? ` (${data.skipped.length} merged skipped)` : '';
    toast(`${data.n_classified} cells → ${label}${skipped}`);
    if (data.stats) updateStats(data.stats);
    // Update cellData locally and refresh the DOM for each classified cell.
    // We do a single /api/cell fetch only for the currently selected cell (to update the sidebar);
    // all other cells get a lightweight local-only refresh based on the action.
    for (const ref of data.classified || []) {
      if (action === 'CLEAR') {
        if (cellData[ref]) { cellData[ref] = { ...cellData[ref], choice: '', name: '', type: '' }; }
      } else {
        const existing = cellData[ref] || {};
        cellData[ref] = { ...existing, choice: action, table_role: '', row_class: '', is_table_end: false };
      }
      refreshCell(ref);
    }
    _clearRange();
    // Refresh the sidebar for the currently focused cell
    if (selectedRef) selectCell(selectedRef);
  } catch (err) {
    toast('Batch classify error: ' + err, true);
  }
  return true;
}

// ── Light navigation: update highlight immediately, defer API fetch ─────────
let _arrowDebounce = null;
function lightMoveToCell(ref) {
  const oldRef = selectedRef;
  selectedRef = ref;
  _refreshCellsByRef([oldRef, ref].filter(Boolean));
  if (gridApi) {
    const [col, rowStr] = _parseColRow(ref);
    if (col) {
      const node = gridApi.getRowNode(rowStr);
      if (node) {
        gridApi.ensureColumnVisible(col);
        gridApi.ensureIndexVisible(node.rowIndex, 'middle');
      }
    }
  }
}

// ── Keyboard shortcuts ─────────────────────────────────────────────────────
function loadKeyboardShortcuts() {
  document.addEventListener('keydown', e => {
    // Ctrl+G → focus jump input (like Cmd+G in spreadsheets)
    if ((e.ctrlKey || e.metaKey) && e.key === 'g') {
      e.preventDefault();
      document.getElementById('cell-jump').focus();
      document.getElementById('cell-jump').select();
      return;
    }
    // Ignore when typing in an input / select / textarea
    if (['INPUT','SELECT','TEXTAREA'].includes(e.target.tagName)) return;
    // Ignore when the mini-table editor modal is open
    if (Alpine?.store?.('gx')?.trmOpen) return;

    if (e.ctrlKey && e.key === 'z') { e.preventDefault(); undoLast(); return; }

    // ── Arrow-key cell navigation (with optional Shift+Arrow range selection) ─
    // lightMoveToCell() updates highlight instantly; selectCell() (API fetch)
    // is debounced so rapid key-repeat doesn't flood the server.
    const arrowDeltas = { ArrowUp: [-1,0], ArrowDown: [1,0], ArrowLeft: [0,-1], ArrowRight: [0,1] };
    if (arrowDeltas[e.key] && selectedRef) {
      e.preventDefault();
      const match = selectedRef.match(/^([A-Z]+)(\d+)$/);
      if (match) {
        const [dr, dc] = arrowDeltas[e.key];
        // Extend from _rangeEnd (or selectedRef if no range yet) when Shift held
        const fromRef = (e.shiftKey && _rangeEnd) ? _rangeEnd : selectedRef;
        const fromMatch = fromRef.match(/^([A-Z]+)(\d+)$/);
        const curCol = colLetterToNum(fromMatch ? fromMatch[1] : match[1]);
        const curRow = parseInt(fromMatch ? fromMatch[2] : match[2]);
        let newRow = curRow + dr, newCol = curCol + dc;
        // Skip merged ghost cells
        for (let attempt = 0; attempt < 20; attempt++) {
          if (newRow < 1 || newCol < 1) break;
          const newRef = numToColLetter(newCol) + newRow;
          const newColLtr = numToColLetter(newCol);
          const cellExists = gridApi
            ? !!gridApi.getRowNode(String(newRow)) && !!gridApi.getColumn(newColLtr)
            : !!document.querySelector(`td[data-ref="${newRef}"]`);
          if (cellExists) {
            if (e.shiftKey && selectedAction === 'T') {
              // In T (table) mode: Shift+Arrow extends the end-ref range.
              // Store the original anchor so openTableModal uses it, then
              // move selectedRef to the cursor so subsequent Shift+Arrow
              // steps accumulate from the right position.
              if (!_tModeAnchor) _tModeAnchor = selectedRef;
              handleShiftClick(newRef);   // updates f-T-end_ref
              // Show the selection rectangle so the user sees the range growing
              _rangeAnchor = _tModeAnchor;
              _rangeEnd    = newRef;
              _renderRangeHighlight();
              lightMoveToCell(newRef);    // advances cursor (selectedRef)
            } else if (e.shiftKey) {
              // General range selection
              if (!_rangeAnchor) _rangeAnchor = selectedRef;
              _rangeEnd = newRef;
              _renderRangeHighlight();
              // Move the focus indicator to the far corner (no API call — stay fast)
              lightMoveToCell(newRef);
            } else {
              // Plain Arrow — cancel any range and T-anchor, move normally
              _clearRange();
              _tModeAnchor = null;
              lightMoveToCell(newRef);
              clearTimeout(_arrowDebounce);
              _arrowDebounce = setTimeout(() => selectCell(selectedRef), 120);
            }
            break;
          }
          newRow += dr; newCol += dc;
        }
      }
      return;
    }

    if (!selectedRef) return;
    const k = e.key.toUpperCase();

    // ── Classification keys: apply to range if one is active ──────────────
    if (['L','V','I'].includes(k) || k === 'DELETE' || k === 'BACKSPACE') {
      const batchAction = k === 'DELETE' || k === 'BACKSPACE' ? 'CLEAR' : k;
      if (_rangeAnchor && _rangeEnd) {
        e.preventDefault();
        _classifyRange(batchAction);
        return;
      }
    }

    // Block L/V/I/T keys when the selected cell is inside a table range
    const _inTable = Alpine?.store?.('gx')?.panelMode === 'table';
    if (_inTable) {
      if (k === 'ENTER') openTableModal();
      else if (k === 'ESCAPE') { _exitTableMode(); prefillForm({}); _clearRange(); }
      return;
    }

    if (k === 'L') selectAction('L');
    else if (k === 'V') selectAction('V');
    else if (k === 'T') selectAction('T');
    else if (k === 'I') selectAction('I');
    else if (k === 'ENTER' && selectedAction) applyClassification();
    else if (k === 'ESCAPE') { cancelClassification(); _clearRange(); }
  });
}

// ── Theme ─────────────────────────────────────────────────────────────────
function toggleTheme() {
  const t = document.documentElement.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
  document.documentElement.setAttribute('data-theme', t);
  // ☀️ when dark (to switch to light), 🌙 when light (to switch to dark)
  document.querySelector('#navbar .btn[onclick="toggleTheme()"]').textContent = t === 'dark' ? '☀️' : '🌙';
}

// ── Utilities ─────────────────────────────────────────────────────────────
function escHtml(s) {
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
function slugify(s) {
  return s.toLowerCase().replace(/[^a-z0-9]+/g,'_').replace(/^_|_$/g,'').substring(0,40) || 'field';
}

/**
 * When classifying a cell as V, look at adjacent cells for an L (label) and
 * derive a suggested field name from it.
 * Direction LR: check columns to the left in the same row (up to 3 back).
 * Direction TD: check rows above in the same column (up to 3 back).
 * Returns slugified label text, or '' if no L neighbour found.
 */
function suggestNameFromLabel(ref) {
  const direction = (document.getElementById('cfg-direction') || {}).value || 'LR';
  const match = ref.match(/^([A-Z]+)(\d+)$/);
  if (!match) return '';
  const colIdx  = colLetterToNum(match[1]);  // 1-based
  const rowIdx  = parseInt(match[2], 10);

  const candidates = [];
  if (direction === 'LR') {
    // Look left: same row, columns colIdx-1 down to max(1, colIdx-3)
    for (let c = colIdx - 1; c >= Math.max(1, colIdx - 3); c--) {
      candidates.push(numToColLetter(c) + rowIdx);
    }
  } else {
    // Look up: same column, rows rowIdx-1 down to max(1, rowIdx-3)
    for (let r = rowIdx - 1; r >= Math.max(1, rowIdx - 3); r--) {
      candidates.push(match[1] + r);
    }
  }

  for (const cref of candidates) {
    const cd = cellData[cref];
    if (cd && cd.choice === 'L') {
      // Prefer lmatch (the anchor text) over raw cell value
      const text = (cd.lmatch || cd.value || cd.raw || '').trim()
                     .replace(/[:\-–—]+$/, '').trim();  // strip trailing : — etc.
      if (text) return slugify(text);
    }
  }
  return '';
}
// ── Table Editor ─────────────────────────────────────────────────────────────

async function _tScanColumns(ref, existingCols) {
  // Fetch auto-detected columns from the anchor row, then render
  try {
    const r = await fetch(`/api/table-scan/${ref}`);
    const data = await r.json();
    const cols = data.columns || [];
    // Merge with any existing user edits (by ref)
    const editMap = {};
    existingCols.forEach(c => { if (c.ref) editMap[c.ref] = c.field_name; });
    const merged = cols.map(c => ({
      ...c,
      field_name: editMap[c.ref] ?? c.suggested_field,
    }));
    _tRenderColumns(merged);
    const hint = document.getElementById('t-col-hint');
    if (hint) hint.style.display = data.columns?.length ? 'none' : 'block';
    document.getElementById('t-col-actions').style.display = 'flex';
    document.getElementById('t-col-header').style.display = 'block';
  } catch(e) {
    console.warn('table-scan failed', e);
  }
}

function _tLoadColumnsFromContext(ctx) {
  // ctx.columns is the stored list: {lbl_name, var_name, var_type, var_match, cell_value}
  const cols = (ctx.columns || []).map(c => {
    const fullName = c.var_name || c.lbl_name || 'IGNORE';
    const tname = ctx.name || '';
    const field_name = (tname && fullName.startsWith(tname + '.'))
      ? fullName.slice(tname.length + 1)
      : fullName;
    return { ref: '', header: c.cell_value || '', field_name, suggested_type: c.var_type || 'string' };
  });
  if (cols.length) {
    _tRenderColumns(cols);
    document.getElementById('t-col-hint').style.display = 'none';
    document.getElementById('t-col-actions').style.display = 'flex';
    document.getElementById('t-col-header').style.display = 'block';
  }
}

function _tRenderColumns(cols) {
  const list = document.getElementById('t-col-list');
  list.innerHTML = '';
  cols.forEach((col, i) => {
    const row = document.createElement('div');
    row.className = 't-col-row';
    row.dataset.ref    = col.ref    || '';
    row.dataset.header = col.header || '';
    row.innerHTML = `
      <span style="font-size:12px;color:var(--text-muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap"
            title="${(col.header || '').replace(/"/g,'&quot;')}">${escHtml(col.header) || '(empty)'}</span>
      <input class="t-col-field" type="text" value="${(col.field_name || '').replace(/"/g,'&quot;')}"
             placeholder="field_name" style="min-width:0">
      <button type="button" class="t-col-del" title="Remove column" onclick="this.closest('.t-col-row').remove()">×</button>
    `;
    list.appendChild(row);
  });
}

function tColAdd() {
  const list = document.getElementById('t-col-list');
  const row = document.createElement('div');
  row.className = 't-col-row';
  row.dataset.ref = ''; row.dataset.header = '';
  row.innerHTML = `
    <input type="text" placeholder="header text" style="font-size:12px;padding:3px 6px;min-width:0">
    <input class="t-col-field" type="text" placeholder="field_name" style="font-size:12px;padding:3px 6px;min-width:0">
    <button type="button" class="t-col-del" onclick="this.closest('.t-col-row').remove()">×</button>
  `;
  list.appendChild(row);
  row.querySelectorAll('input')[0].addEventListener('input', function() {
    row.dataset.header = this.value;
    const fieldEl = row.querySelector('.t-col-field');
    if (!fieldEl.value) fieldEl.value = slugify(this.value);
  });
}

function tColRescan() {
  // Re-detect columns but preserve edits the user has made
  const existingCols = Array.from(document.querySelectorAll('#t-col-list .t-col-row')).map(r => ({
    ref: r.dataset.ref, field_name: r.querySelector('.t-col-field')?.value || '',
  }));
  if (selectedRef) _tScanColumns(selectedRef, existingCols);
}

// Show a brief status line in the sidebar after row_configs are saved
function _tUpdateRowConfigStatus(ctx) {
  const statusEl = document.getElementById('t-row-config-status');
  if (!statusEl) return;
  const wrc = ctx._web_row_configs || [];
  if (!wrc.length) { statusEl.style.display = 'none'; return; }
  const nH = wrc.filter(r=>r.row_type==='header').length;
  const nD = wrc.filter(r=>r.row_type==='data'||r.row_type==='data_inherited').length;
  const nF = wrc.filter(r=>r.row_type==='footer').length;
  const nS = wrc.filter(r=>r.row_type==='skip').length;
  statusEl.textContent = `✓ ${wrc.length} rows: ${nH}H ${nD}D ${nF}F ${nS}Skip`;
  statusEl.style.display = 'block';
}

// After classifying a T cell, refresh all the T/T-HEAD/T-DATA cells on the grid
async function _tRefreshTableCells(anchorRef) {
  try {
    const r   = await fetch('/api/state');
    const data = await r.json();
    const cells = data.cells || [];
    for (const c of cells) {
      const meta = c;
      if (meta.choice === 'T-HEAD' || meta.choice === 'T-DATA' || meta.choice === 'T') {
        // just trigger a visual refresh for that cell
        refreshCell(meta.ref);
      }
    }
  } catch(e) {}
  // Also refresh via loadSheet for full grid re-render
  await loadSheet();
}

// ── Value type change: hide match controls for image ──────────────────────────
function onVTypeChange(type) {
  const matchModeRow = document.getElementById('f-V-match_mode')?.closest('.form-group');
  const matchRow     = document.getElementById('f-V-match')?.closest('.form-group');
  const hide = type === 'image';
  if (matchModeRow) matchModeRow.style.display = hide ? 'none' : '';
  if (matchRow)     matchRow.style.display     = hide ? 'none' : '';
}

// ── Modifier checkbox dropdown ────────────────────────────────────────────────
function toggleModDropdown(e) {
  e && e.stopPropagation();
  const trigger = document.getElementById('mod-trigger');
  const menu    = document.getElementById('mod-menu');
  if (!trigger || !menu) return;

  // Move to body so it escapes #panel's overflow:hidden (Chrome clips
  // position:fixed inside overflow:hidden flex containers).
  if (menu.parentElement !== document.body) {
    document.body.appendChild(menu);
  }

  const isOpen = menu.classList.contains('open');
  if (isOpen) {
    menu.classList.remove('open');
  } else {
    const rect = trigger.getBoundingClientRect();
    const menuMinW = 190;  // wide enough for "trim-whitespace"
    menu.style.top      = (rect.bottom + 2) + 'px';
    menu.style.minWidth = Math.max(menuMinW, rect.width) + 'px';
    menu.style.width    = '';
    const leftPos = Math.min(rect.left, window.innerWidth - menuMinW - 8);
    menu.style.left = Math.max(0, leftPos) + 'px';
    menu.classList.add('open');
  }
}
document.addEventListener('click', function(e) {
  const trigger = document.getElementById('mod-trigger');
  const menu    = document.getElementById('mod-menu');
  if (menu && menu.classList.contains('open')) {
    if ((!trigger || !trigger.contains(e.target)) && !menu.contains(e.target)) {
      menu.classList.remove('open');
    }
  }
});
// Close the menu if the sidebar scrolls (trigger moves, menu stays fixed otherwise)
document.getElementById('panel-body') && document.getElementById('panel-body').addEventListener('scroll', function() {
  const menu = document.getElementById('mod-menu');
  if (menu) menu.classList.remove('open');
}, { passive: true });

function onModChange(which) {
  const noneEl  = document.getElementById('mod-none');
  const nullEl  = document.getElementById('mod-nullable');
  const nnullEl = document.getElementById('mod-not-null');
  const trimEl  = document.getElementById('mod-trim');
  if (which === 'none') {
    if (noneEl.checked) {
      nullEl.checked = false; nnullEl.checked = false; trimEl.checked = false;
    } else {
      // unchecking none with nothing else on → re-check none
      if (!nullEl.checked && !nnullEl.checked && !trimEl.checked) noneEl.checked = true;
    }
  } else {
    noneEl.checked = false;
    if (which === 'nullable'  && nullEl.checked)  nnullEl.checked = false;
    if (which === 'not-null'  && nnullEl.checked) nullEl.checked  = false;
    // if everything was unchecked, restore none
    if (!nullEl.checked && !nnullEl.checked && !trimEl.checked) noneEl.checked = true;
  }
  _syncModifiers();
}

function _syncModifiers() {
  const isNone  = document.getElementById('mod-none').checked;
  const isNull  = document.getElementById('mod-nullable').checked;
  const isNnull = document.getElementById('mod-not-null').checked;
  const isTrim  = document.getElementById('mod-trim').checked;
  let val, display;
  if (isNone || (!isNull && !isNnull && !isTrim)) {
    val = 'none'; display = 'none';
  } else {
    const parts = [];
    if (isNull)  parts.push('nullable');
    if (isNnull) parts.push('not-null');
    if (isTrim)  parts.push('trim-whitespace');
    val = parts.join(':');
    display = parts.join(' + ');
  }
  document.getElementById('f-V-modifiers').value = val;
  document.getElementById('mod-display').textContent = display;
}

function _applyModifiers(val) {
  val = val || 'none';
  const parts = (val === 'none') ? [] : val.split(':');
  document.getElementById('mod-none').checked    = parts.length === 0;
  document.getElementById('mod-nullable').checked = parts.includes('nullable');
  document.getElementById('mod-not-null').checked  = parts.includes('not-null');
  document.getElementById('mod-trim').checked      = parts.includes('trim-whitespace');
  _syncModifiers();
}

function setVal(id, v) {
  if (id === 'f-V-modifiers') { _applyModifiers(v); return; }
  const el = document.getElementById(id);
  if (el) el.value = v;
}
function setChk(id, v) {
  const el = document.getElementById(id);
  if (el) el.checked = !!v;
}
function setStatus(msg) { document.getElementById('status-msg').textContent = msg; }
function toast(msg, err, durationMs) {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.style.background = err ? '#dc3545' : (msg.startsWith('⚠️') ? '#8a6d00' : '#333');
  el.classList.add('show');
  setTimeout(() => el.classList.remove('show'), durationMs || 2800);
}

// ── Shift+click → set end_ref ─────────────────────────────────────────────
document.addEventListener('keydown', e => {
  if (Alpine?.store?.('gx')?.trmOpen) return;
  if (e.key === 'Shift') document.body.classList.add('shift-mode');
});
document.addEventListener('keyup', e => { if (e.key === 'Shift') document.body.classList.remove('shift-mode'); });

function handleShiftClick(ref) {
  if (selectedAction !== 'T') return false;
  const el = document.getElementById('f-T-end_ref');
  if (!el) return false;
  el.value = ref.toUpperCase();
  el.classList.add('flash');
  setTimeout(() => el.classList.remove('flash'), 800);
  toast(`End ref set to ${ref.toUpperCase()}`);
  return true;  // consumed
}

// ── T-mode keyboard anchor tracking ──────────────────────────────────────
// When the user uses Shift+Arrow in T mode to extend the end-ref range,
// selectedRef moves to the cursor position so subsequent Shift+Arrow steps
// accumulate correctly.  But openTableModal() must use the original anchor
// (the cell that had T selected), not the cursor.  _tModeAnchor captures
// that original anchor; openTableModal() prefers it over selectedRef.
// It is cleared whenever the user makes a plain click / plain Arrow move.
let _tModeAnchor = null;

// ── Table Row Editor Modal ─────────────────────────────────────────────────

let _trmAnchorRef   = '';   // anchor (top-left) ref being edited
let _trmRowCount    = 0;    // sequential id counter for synthetic rows
let _trmHeaderNames = [];   // header col names collected at grid render (index = col index)

function _trmSlugify(s) {
  return String(s).toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/^_|_$/g, '') || 'col';
}

/** Live-update the field preview hint when the user edits a data-row name input. */
function _trmUpdateNameHint(input) {
  const hint = input.closest('.trm-col-ed')?.querySelector('.trm-name-hint-field');
  if (hint) hint.textContent = input.value.trim() || '…';
}

/** Refresh all .trm-name-hint-tbl spans when the table name changes.
 *  Called by the trm-name input's oninput handler.
 */
function _trmRefreshTableNameHints() {
  const name = (document.getElementById('trm-name')?.value || '').trim() || '…';
  document.querySelectorAll('.trm-name-hint-tbl').forEach(el => {
    el.textContent = name;
    el.dataset.tbl = name;
  });
}

// Ref ↔ (row, col) helpers for JS
function _parseRef(ref) {
  const m = String(ref).match(/^([A-Za-z]+)(\d+)$/);
  if (!m) return [0, 0];
  const colStr = m[1].toUpperCase();
  let col = 0;
  for (const ch of colStr) col = col * 26 + (ch.charCodeAt(0) - 64);
  return [parseInt(m[2], 10), col];
}
function _buildRef(row, col) {
  let s = '';
  while (col > 0) { const r = (col - 1) % 26; s = String.fromCharCode(65 + r) + s; col = Math.floor((col - 1) / 26); }
  return s + row;
}

/**
 * Scan cellData for all T / T-HEAD / T-DATA cells belonging to anchorRef and
 * return the bottom-right cell ref.  Used to auto-determine end_ref when the
 * user hasn't shift-clicked yet (e.g. after loading a preloaded pattern).
 */
function _inferTableEndRef(anchorRef) {
  let maxRow = 0, maxCol = 0;
  Object.values(cellData).forEach(info => {
    const ref = info.ref || '';
    const choice = info.choice || '';
    const isAnchor  = ref === anchorRef && choice === 'T';
    const isMember  = (choice === 'T-HEAD' || choice === 'T-DATA') && info.anchor === anchorRef;
    if (!isAnchor && !isMember) return;
    const [r, c] = _parseRef(ref);
    if (r > maxRow) maxRow = r;
    if (c > maxCol) maxCol = c;
  });
  return (maxRow > 0 && maxCol > 0) ? _buildRef(maxRow, maxCol) : '';
}
const _TRM_ROW_TYPES = [
  { val: 'header',         label: 'Header:N',    hasN: true,  hasCols: true  },
  { val: 'data',           label: 'Data:*',       hasN: false, hasCols: true  },
  { val: 'data_inherited', label: 'Data (inherit)',hasN: false, hasCols: false },
  { val: 'footer',         label: 'Footer:N',     hasN: true,  hasCols: true  },
  { val: 'skip',           label: 'Skip',         hasN: false, hasCols: true,  isSkip: true },
  { val: 'ignore',         label: 'Ignore',       hasN: false, hasCols: false },
];
// V=Variable, L=Label, I=Ignore, E=Empty  (C was removed — had no distinct table role)
const _TRM_ROLES       = ['V','L','I','E'];
const _TRM_ROLE_LABELS = {V:'V — Variable', L:'L — Label', I:'I — Ignore', E:'E — Empty'};
const _TRM_TYPES = {
  basic:    ['string', 'number', 'date', 'boolean', 'url', 'image'],
  advanced: ['integer', 'currency', 'datetime', 'time'],
};
function _trmTypeOptions(selected) {
  return [
    { label: 'Basic',    types: _TRM_TYPES.basic },
    { label: 'Advanced', types: _TRM_TYPES.advanced },
  ].map(g =>
    `<optgroup label="${g.label}">${g.types.map(t =>
      `<option value="${t}"${t === selected ? ' selected' : ''}>${t}</option>`
    ).join('')}</optgroup>`
  ).join('');
}
const _TRM_SKIP_COND = ['IGNORE','EMPTY','label'];

async function openTableModal() {
  const nameEl   = document.getElementById('f-T-name');
  const multEl   = document.getElementById('f-T-mult');
  const endEl    = document.getElementById('f-T-end_ref');
  // _tModeAnchor is set when Shift+Arrow was used to extend the T range;
  // in that case selectedRef has moved to the cursor (end-ref position),
  // so we must use the stored anchor instead.
  const anchorRef = _tModeAnchor || selectedRef || '';

  if (!anchorRef) { toast('Select the top-left table cell first', true); return; }

  // Resolve end_ref: sidebar input → inferred from visible T-HEAD/T-DATA cells → error
  let endRef = (endEl?.value || '').toUpperCase().trim();
  if (!endRef) {
    endRef = _inferTableEndRef(anchorRef);
    if (endRef && endEl) { endEl.value = endRef; }  // write back so sidebar shows it
  }
  if (!endRef && _rangeAnchor && _rangeEnd &&
      _rangeAnchor.toUpperCase() === anchorRef.toUpperCase()) {
    endRef = _rangeEnd.toUpperCase();
    if (endEl) { endEl.value = endRef; }
  }
  if (!endRef) {
    toast('Set the bottom-right cell (end ref): Shift+click any cell in the table', true);
    return;
  }

  _trmAnchorRef = anchorRef;

  // Fetch range data
  let rangeData;
  try {
    const r = await fetch(`/api/table-range/${anchorRef}/${endRef}`);
    if (!r.ok) { toast('Could not fetch table range — check the end ref', true); return; }
    rangeData = await r.json();
  } catch(e) { toast('Network error fetching range', true); return; }

  // Populate settings
  document.getElementById('trm-name').value = nameEl?.value || '';
  const multVal = multEl?.value || '*';
  const multSt = document.getElementById('trm-mult');
  const multCi = document.getElementById('trm-mult-custom');
  if (['*','1','2','3'].includes(multVal)) {
    multSt.value = multVal; multCi.hidden = true;
  } else {
    multSt.value = '{n,m}'; multCi.value = multVal; multCi.hidden = false;
  }
  document.getElementById('trm-range-display').textContent = `${anchorRef} → ${endRef}`;

  // Try to load existing row_configs for editing
  let existingConfigs = null;
  try {
    const cx = await fetch(`/api/table-context/${anchorRef}`);
    if (cx.ok) {
      const ctx = await cx.json();
      if (ctx._web_row_configs?.length) existingConfigs = ctx._web_row_configs;
    }
  } catch(e) {}

  // Render the grid
  _trmRenderGrid(rangeData, existingConfigs);

  // Open modal
  Alpine.store('gx').trmOpen = true;
}

function closeTableModal() {
  Alpine.store('gx').trmOpen = false;
  // Drop shift-select mode so Shift+click no longer fires handleShiftClick
  // after the modal is dismissed.  applyTableModal → loadSheet resets state
  // automatically; Cancel without Apply must do it explicitly.
  if (selectedAction === 'T') {
    selectedAction = null;
    document.querySelectorAll('.cbtn').forEach(b => b.classList.remove('active'));
  }
}

// ── Mini-table editor state ──────────────────────────────────────────────────
let _trmState = { rows: [], colRefs: [], rawCells: [], inferredTypes: [] };
let _trmSelectedRow = -1;
let _trmSelectedCol = -1;  // -1 = row-type ctrl cell selected

function _trmRowClass(rowType) {
  return {header:'H', data:'D', data_inherited:'Di', footer:'F', skip:'S', ignore:'X'}[rowType] || 'X';
}

function _trmInitState(rangeData, existingConfigs) {
  const rows = rangeData.rows || [];
  const cfgByRow = {};
  if (existingConfigs) existingConfigs.forEach(rc => { cfgByRow[rc.sheet_row] = rc; });

  _trmHeaderNames = [];
  if (rows[0]) {
    const hCfg = cfgByRow[rows[0].sheet_row] || null;
    rows[0].cells.forEach((cell, ci) => {
      _trmHeaderNames.push(hCfg?.cols?.[ci]?.name || _trmSlugify(String(cell.raw ?? '')));
    });
  }

  _trmState.colRefs      = rows[0] ? rows[0].cells.map(c => c.ref) : [];
  _trmState.rawCells     = rows.map(r => r.cells.map(c => c.raw ?? ''));
  _trmState.inferredTypes= rows.map(r => r.cells.map(c => c.inferred_type || 'string'));

  _trmState.rows = rows.map((row, ri) => {
    const sheetRow  = row.sheet_row || 0;
    const cfg       = cfgByRow[sheetRow] || null;
    const defaultType = ri === 0 ? 'header' : (ri === 1 ? 'data' : 'data_inherited');
    const rowType   = cfg?.row_type || defaultType;

    if (rowType === 'ignore' || rowType === 'data_inherited') {
      return { sheet_row: sheetRow, row_type: rowType };
    }

    const cols = (cfg?.cols || row.cells.map((cell, ci) => {
      const defRole = rowType === 'data' ? 'V' : 'L';
      return {
        role:        defRole,
        name:        rowType === 'data' ? (_trmHeaderNames[ci] || '') : (_trmSlugify(String(cell.raw ?? '')) + (rowType === 'header' ? '_header_label' : '_label')),
        ftype:       (cell.inferred_type && cell.inferred_type !== 'string') ? cell.inferred_type : 'string',
        match:       '.*',
        modifiers:   'none',
        lmatch:      String(cell.raw ?? ''),
        lmatch_mode: 'literal',
        condition:   'IGNORE',
      };
    })).map((c, ci) => ({
      role:        c.role        || (rowType === 'data' ? 'V' : 'L'),
      name:        c.name        ?? '',
      ftype:       c.ftype       || 'string',
      match:       c.match       || '.*',
      modifiers:   c.modifiers   || 'none',
      lmatch:      c.lmatch      ?? (row.cells[ci] ? String(row.cells[ci].raw ?? '') : ''),
      lmatch_mode: c.lmatch_mode || 'literal',
      condition:   c.condition   || 'IGNORE',
    }));

    const stateRow = { sheet_row: sheetRow, row_type: rowType, cols };
    if (rowType === 'header' || rowType === 'footer') stateRow.row_n = cfg?.row_n || 1;
    return stateRow;
  });
}

function _trmRenderGrid(rangeData, existingConfigs) {
  _trmSelectedRow = -1;
  _trmSelectedCol = -1;
  _trmInitState(rangeData, existingConfigs);
  _trmRenderGridDOM();
  _trmRenderPanel();
}

function _trmRenderGridDOM() {
  const tbody = document.getElementById('trm-tbody');
  tbody.innerHTML = '';
  const nCols = _trmState.colRefs.length;

  // Column ref header row
  const headTr = document.createElement('tr');
  headTr.innerHTML = `<th style="font-weight:700;padding:4px 8px;min-width:96px;background:var(--surface-raised,var(--surface))">Row type</th>`;
  _trmState.colRefs.forEach(ref => {
    headTr.innerHTML += `<th style="font-weight:700;padding:4px 8px;min-width:100px">${ref}</th>`;
  });
  tbody.appendChild(headTr);

  _trmState.rows.forEach((row, ri) => {
    const tr = document.createElement('tr');
    tr.className = `trm-row-${_trmRowClass(row.row_type)}`;
    tr.dataset.rowIdx = ri;

    const ctrlTd = document.createElement('td');
    ctrlTd.className = 'trm-row-ctrl';
    ctrlTd.dataset.rowIdx = ri;
    ctrlTd.onclick = () => _trmClickRowType(ri);
    ctrlTd.innerHTML = _trmRowCtrlHTML(row);
    tr.appendChild(ctrlTd);

    const rawRow = _trmState.rawCells[ri] || [];
    for (let ci = 0; ci < nCols; ci++) {
      const td = document.createElement('td');
      td.className = 'trm-cell';
      td.dataset.rowIdx = ri;
      td.dataset.colIdx = ci;
      td.onclick = () => _trmClickCell(ri, ci);
      const raw = rawRow[ci] ?? '';
      td.innerHTML = `<div class="trm-cell-raw" title="${_esc(String(raw))}">${_esc(String(raw))}</div>${_trmColBadgeHTML(row, ci)}`;
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  });
}

function _trmRowCtrlHTML(row) {
  const colorMap = { H:'#7c3aed', D:'#0284c7', Di:'#93c5fd', F:'#0891b2', S:'#d97706', X:'#9ca3af' };
  const col = colorMap[_trmRowClass(row.row_type)] || '#9ca3af';
  const label = _TRM_ROW_TYPES.find(r => r.val === row.row_type)?.label || row.row_type;
  let html = `<span class="trm-rt-chip" style="color:${col};border-color:${col}">${label}</span>`;
  if (row.row_type === 'header' || row.row_type === 'footer') {
    html += `<div style="font-size:10px;color:var(--text-muted);margin-top:3px"># ${row.row_n || 1}</div>`;
  }
  return html;
}

function _trmColBadgeHTML(row, ci) {
  const rt = row.row_type;
  if (rt === 'data_inherited') return `<span class="trm-cell-badge-inh">↑ inherits</span>`;
  if (rt === 'ignore')         return `<span class="trm-cell-badge-inh">—</span>`;
  if (rt === 'skip') {
    const cond = row.cols?.[ci]?.condition || 'IGNORE';
    return `<span class="trm-cell-badge trm-cell-badge-I">${cond}</span>`;
  }
  const col  = row.cols?.[ci] || {};
  const role = col.role || (rt === 'data' ? 'V' : 'L');
  const name = col.name || '';
  return `<span class="trm-cell-badge trm-cell-badge-${role}">${role}</span>${name ? `<div style="font-size:13px;color:var(--text-muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${_esc(name)}</div>` : ''}`;
}

function _trmClickRowType(ri) {
  _trmSelectedRow = ri;
  _trmSelectedCol = -1;
  _trmUpdateSelectionHighlight();
  _trmRenderPanel();
}

function _trmClickCell(ri, ci) {
  _trmSelectedRow = ri;
  _trmSelectedCol = ci;
  _trmUpdateSelectionHighlight();
  _trmRenderPanel();
}

function _trmUpdateSelectionHighlight() {
  document.querySelectorAll('#trm-tbody .trm-sel').forEach(el => el.classList.remove('trm-sel'));
  if (_trmSelectedRow === -1) return;
  if (_trmSelectedCol === -1) {
    const ctrl = document.querySelector(`#trm-tbody .trm-row-ctrl[data-row-idx="${_trmSelectedRow}"]`);
    if (ctrl) ctrl.classList.add('trm-sel');
  } else {
    const cell = document.querySelector(`#trm-tbody .trm-cell[data-row-idx="${_trmSelectedRow}"][data-col-idx="${_trmSelectedCol}"]`);
    if (cell) cell.classList.add('trm-sel');
  }
}

function _trmRenderPanel() {
  const content = document.getElementById('trm-panel-content');
  const empty   = document.getElementById('trm-panel-empty');
  if (!content || !empty) return;
  if (_trmSelectedRow === -1) {
    content.hidden = true; empty.style.display = '';
    return;
  }
  empty.style.display = 'none'; content.hidden = false;
  const row = _trmState.rows[_trmSelectedRow];
  if (!row) return;
  content.innerHTML = _trmSelectedCol === -1 ? _trmPanelRowTypeHTML(row) : _trmPanelCellHTML(row, _trmSelectedCol);
}

function _trmPanelRowTypeHTML(row) {
  const rt   = row.row_type;
  const isHF = rt === 'header' || rt === 'footer';
  const btns = _TRM_ROW_TYPES.map(r =>
    `<button class="trm-panel-rt-btn${rt===r.val?' active':''}" onclick="_trmPanelRowTypeClick('${r.val}')">${r.label}</button>`
  ).join('');
  return `<div class="trm-panel-section-title">Row type</div>
    <div class="trm-panel-rowtype-btns">${btns}</div>
    <div class="trm-panel-n-row" style="display:${isHF?'flex':'none'}" id="trm-panel-n-row">
      <label for="trm-panel-row-n"># (row number):</label>
      <input type="number" id="trm-panel-row-n" min="1" value="${row.row_n||1}"
             oninput="_trmPanelRowNChange(this.value)">
    </div>`;
}

function _trmPanelCellHTML(row, ci) {
  const ref      = _trmState.colRefs[ci] || `col${ci+1}`;
  const cellRaw  = _trmState.rawCells[_trmSelectedRow]?.[ci] ?? '';
  const cellType = _trmState.inferredTypes[_trmSelectedRow]?.[ci] ?? 'string';
  const colorMap = { H:'#7c3aed', D:'#0284c7', Di:'#93c5fd', F:'#0891b2', S:'#d97706', X:'#9ca3af' };
  const col      = colorMap[_trmRowClass(row.row_type)] || '#9ca3af';
  const rowLabel = _TRM_ROW_TYPES.find(r => r.val === row.row_type)?.label || row.row_type;
  let html = `<div class="trm-panel-cell-title">
      <span class="trm-panel-ref">${ref}</span>
      <span class="trm-leg-chip" style="color:${col};background:${col}1a;border:1px solid ${col}4d">${rowLabel}</span>
    </div>`;
  const rt = row.row_type;
  if (rt === 'data_inherited' || rt === 'ignore') {
    html += `<div class="trm-panel-inherit">${rt==='data_inherited'?'↑ Inherits schema from the data row':'— Row is ignored'}</div>`;
  } else if (rt === 'skip') {
    const colCfg = row.cols?.[ci] || {};
    const cond   = colCfg.condition || 'IGNORE';
    const lmatch = colCfg.lmatch || '';
    html += `<div class="form-group"><label>Skip condition</label>
      <select onchange="_trmPanelSkipCondChange(this.value,${ci})">
        ${['IGNORE','EMPTY','label'].map(c=>`<option value="${c}"${c===cond?' selected':''}>${c}</option>`).join('')}
      </select></div>
    <div class="form-group" id="trm-skip-lmatch-group" style="display:${cond==='label'?'block':'none'}">
      <label>Match text</label>
      <input type="text" value="${_esc(lmatch)}" placeholder="label text"
             oninput="_trmPanelFieldChange('lmatch',this.value,${ci})">
    </div>`;
  } else {
    const colCfg = row.cols?.[ci] || {};
    const role   = colCfg.role || (rt==='data'?'V':'L');
    html += _trmPanelRoleFieldsHTML(role, rt, ci, colCfg);
  }
  html += `<div class="trm-panel-cell-info">
    <div class="trm-panel-cell-info-title">Cell info</div>
    <div><span class="trm-panel-ci-label">Value</span><span class="trm-panel-ci-val">${_esc(String(cellRaw))}</span></div>
    <div><span class="trm-panel-ci-label">Type</span><span class="badge-type">${cellType}</span></div>
    <div><span class="trm-panel-ci-label">Ref</span><span class="trm-panel-ci-val" style="font-family:monospace">${ref}</span></div>
  </div>`;
  return html;
}

function _trmPanelRoleFieldsHTML(role, rt, ci, col) {
  const roleLabels = {L:'Label (L)',V:'Value (V)',I:'Ignore (I)',E:'Empty (E)'};
  let html = `<div class="trm-panel-role-btns">
    ${['L','V','I','E'].map(r=>`<button class="trm-panel-role-btn trm-role-${r}${role===r?' active':''}"
        onclick="_trmPanelRoleClick('${r}',${ci})">${roleLabels[r]}</button>`).join('')}
  </div>`;
  if (role === 'I') {
    html += `<div class="trm-panel-inherit">Column will be ignored during extraction</div>`;
  } else if (role === 'E') {
    html += `<div class="trm-panel-inherit">Column expected to be blank</div>`;
  } else if (role === 'L') {
    const name        = col.name        || '';
    const lmatch      = col.lmatch      || '';
    const lmatch_mode = col.lmatch_mode || 'literal';
    html += `<div class="form-group"><label>Label name</label>
      <input type="text" value="${_esc(name)}" placeholder="e.g. invoice_no_label"
             oninput="_trmPanelFieldChange('name',this.value,${ci})"></div>
    <div class="form-group"><label>Match mode</label>
      <select onchange="_trmPanelFieldChange('lmatch_mode',this.value,${ci})">
        <option value="literal"${lmatch_mode==='literal'?' selected':''}>literal (default)</option>
        <option value="glob"${lmatch_mode==='glob'?' selected':''}>glob</option>
        <option value="regexp"${lmatch_mode==='regexp'?' selected':''}>regexp</option>
      </select></div>
    <div class="form-group"><label>Match text / pattern</label>
      <input type="text" value="${_esc(lmatch)}" placeholder="exact text or pattern"
             oninput="_trmPanelFieldChange('lmatch',this.value,${ci})"></div>`;
  } else {
    const name  = col.name     || (_trmHeaderNames[ci] || '');
    const ftype = col.ftype    || 'string';
    const match = col.match    || '.*';
    const mods  = col.modifiers|| 'none';
    const tblName = (document.getElementById('trm-name')?.value || '').trim() || '…';
    html += `<div class="form-group"><label>Field name</label>
      <input type="text" id="trm-panel-fname" value="${_esc(name)}" placeholder="e.g. total_amount"
             oninput="_trmPanelFieldChange('name',this.value,${ci})">
      <div class="trm-panel-name-hint">
        <span class="trm-name-hint-tbl">${_esc(tblName)}</span>.<span id="trm-panel-fname-preview">${_esc(name||'…')}</span>
      </div></div>
    <div class="form-group"><label>Type</label>
      <select onchange="_trmPanelFieldChange('ftype',this.value,${ci})">${_trmTypeOptions(ftype)}</select></div>
    <div class="form-group"><label>Match pattern</label>
      <input type="text" value="${_esc(match)}" placeholder=".*"
             oninput="_trmPanelFieldChange('match',this.value,${ci})"></div>
    <div class="form-group"><label>Modifiers</label>
      <select onchange="_trmPanelFieldChange('modifiers',this.value,${ci})">
        <option value="none"${mods==='none'?' selected':''}>none</option>
        <option value="nullable"${mods==='nullable'?' selected':''}>nullable</option>
        <option value="not-null"${mods==='not-null'?' selected':''}>not-null</option>
        <option value="trim-whitespace"${mods==='trim-whitespace'?' selected':''}>trim-whitespace</option>
        <option value="nullable:trim-whitespace"${mods==='nullable:trim-whitespace'?' selected':''}>nullable + trim</option>
      </select></div>`;
  }
  return html;
}

function _trmPanelRowTypeClick(val) {
  const row = _trmState.rows[_trmSelectedRow];
  if (!row) return;
  row.row_type = val;
  const isHF = val === 'header' || val === 'footer';
  if (isHF && !row.row_n) row.row_n = 1;
  if (val === 'ignore' || val === 'data_inherited') {
    delete row.cols;
  } else if (!row.cols) {
    const rawRow = _trmState.rawCells[_trmSelectedRow] || [];
    row.cols = _trmState.colRefs.map((_, ci) => ({
      role: val==='data'?'V':'L',
      name: val==='data'?(_trmHeaderNames[ci]||''):(_trmSlugify(String(rawRow[ci]??''))+(val==='header'?'_header_label':'_label')),
      ftype:'string', match:'.*', modifiers:'none',
      lmatch:String(rawRow[ci]??''), lmatch_mode:'literal', condition:'IGNORE',
    }));
  }
  const tr = document.querySelector(`#trm-tbody tr[data-row-idx="${_trmSelectedRow}"]`);
  if (tr) {
    tr.className = `trm-row-${_trmRowClass(val)}`;
    const ctrlTd = tr.querySelector('.trm-row-ctrl');
    if (ctrlTd) { ctrlTd.innerHTML = _trmRowCtrlHTML(row); ctrlTd.classList.add('trm-sel'); }
    for (let ci = 0; ci < _trmState.colRefs.length; ci++) _trmUpdateCellBadge(_trmSelectedRow, ci);
  }
  _trmRenderPanel();
}

function _trmPanelRowNChange(val) {
  const row = _trmState.rows[_trmSelectedRow];
  if (row) row.row_n = parseInt(val, 10) || 1;
  const ctrl = document.querySelector(`#trm-tbody .trm-row-ctrl[data-row-idx="${_trmSelectedRow}"]`);
  if (ctrl && row) ctrl.innerHTML = _trmRowCtrlHTML(row);
}

function _trmPanelRoleClick(role, ci) {
  const row = _trmState.rows[_trmSelectedRow];
  if (!row?.cols) return;
  if (!row.cols[ci]) row.cols[ci] = {};
  row.cols[ci].role = role;
  _trmUpdateCellBadge(_trmSelectedRow, ci);
  _trmRenderPanel();
}

function _trmPanelFieldChange(field, val, ci) {
  const row = _trmState.rows[_trmSelectedRow];
  if (!row?.cols) return;
  if (!row.cols[ci]) row.cols[ci] = {};
  row.cols[ci][field] = val;
  if (field === 'name') {
    const p = document.getElementById('trm-panel-fname-preview');
    if (p) p.textContent = val || '…';
    _trmUpdateCellBadge(_trmSelectedRow, ci);
  }
}

function _trmPanelSkipCondChange(val, ci) {
  const row = _trmState.rows[_trmSelectedRow];
  if (!row?.cols) return;
  if (!row.cols[ci]) row.cols[ci] = {};
  row.cols[ci].condition = val;
  const lg = document.getElementById('trm-skip-lmatch-group');
  if (lg) lg.style.display = val === 'label' ? 'block' : 'none';
  _trmUpdateCellBadge(_trmSelectedRow, ci);
}

function _trmUpdateCellBadge(ri, ci) {
  const td = document.querySelector(`#trm-tbody .trm-cell[data-row-idx="${ri}"][data-col-idx="${ci}"]`);
  if (!td) return;
  const row = _trmState.rows[ri];
  if (!row) return;
  const rawDiv = td.querySelector('.trm-cell-raw');
  td.innerHTML = (rawDiv ? rawDiv.outerHTML : '') + _trmColBadgeHTML(row, ci);
  if (ri === _trmSelectedRow && ci === _trmSelectedCol) td.classList.add('trm-sel');
}

function _esc(s) { return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }

function trmAddRow(rowType) {
  const lastRow   = _trmState.rows.at(-1);
  const lastSheet = lastRow ? lastRow.sheet_row + 1 : 99;
  const nCols     = _trmState.colRefs.length;
  _trmState.rawCells.push(Array.from({length: nCols}, () => ''));
  _trmState.inferredTypes.push(Array.from({length: nCols}, () => 'string'));
  if (rowType === 'ignore' || rowType === 'data_inherited') {
    _trmState.rows.push({ sheet_row: lastSheet, row_type: rowType });
  } else {
    const cols = Array.from({length: nCols}, (_, ci) => ({
      role: rowType==='data'?'V':'L', name: rowType==='data'?(_trmHeaderNames[ci]||''):'',
      ftype:'string', match:'.*', modifiers:'none',
      lmatch:'', lmatch_mode:'literal', condition:'IGNORE',
    }));
    const newRow = { sheet_row: lastSheet, row_type: rowType, cols };
    if (rowType === 'header' || rowType === 'footer') newRow.row_n = 1;
    _trmState.rows.push(newRow);
  }
  _trmRenderGridDOM();
  _trmUpdateSelectionHighlight();
}

function collectTableRowConfigs() {
  return _trmState.rows.map(row => {
    const rc = { sheet_row: row.sheet_row, row_type: row.row_type };
    if (row.row_type === 'header' || row.row_type === 'footer') rc.row_n = row.row_n || 1;
    if (row.row_type === 'ignore' || row.row_type === 'data_inherited') return rc;
    rc.cols = (row.cols || []).map(col => {
      if (row.row_type === 'skip') {
        return { condition: col.condition||'IGNORE', lmatch: col.condition==='label'?(col.lmatch||''):'' };
      }
      if (col.role === 'I' || col.role === 'E') return { role: col.role, name: col.name||'' };
      if (col.role === 'L') {
        return { role:'L', name:col.name||'', lmatch:col.lmatch||'', lmatch_mode:col.lmatch_mode||'literal', ftype:col.ftype||'string' };
      }
      const out = { role:'V', name:col.name||'', ftype:col.ftype||'string', match:col.match||'.*' };
      if (col.modifiers && col.modifiers !== 'none') out.modifiers = col.modifiers;
      return out;
    });
    return rc;
  });
}

async function applyTableModal() {
  const anchorRef = _trmAnchorRef;
  const name      = document.getElementById('trm-name').value.trim() || 'table';
  const multSel   = document.getElementById('trm-mult').value || '*';
  const mult      = multSel === '{n,m}'
    ? (document.getElementById('trm-mult-custom').value.trim() || '*')
    : multSel;
  const endRefEl  = document.getElementById('f-T-end_ref');
  const endRef    = (endRefEl?.value || '').toUpperCase().trim();
  if (!endRef) { toast('End ref is missing — fill in the bottom-right cell', true); return; }

  const rowConfigs = collectTableRowConfigs();

  try {
    const data = await classify(anchorRef, 'T', {
      name, mult, end_ref: endRef, row_configs: rowConfigs,
    });
    if (data.ok) {
      // Update sidebar fields
      if (document.getElementById('f-T-name')) document.getElementById('f-T-name').value = name;
      if (document.getElementById('f-T-mult')) document.getElementById('f-T-mult').value = mult;
      // Show row config status
      const statusEl = document.getElementById('t-row-config-status');
      if (statusEl) {
        const nH = rowConfigs.filter(r=>r.row_type==='header').length;
        const nD = rowConfigs.filter(r=>r.row_type==='data'||r.row_type==='data_inherited').length;
        const nS = rowConfigs.filter(r=>r.row_type==='skip').length;
        const nF = rowConfigs.filter(r=>r.row_type==='footer').length;
        statusEl.textContent = `✓ ${nH}H ${nD}D ${nF}F ${nS}Skip rows configured`;
        statusEl.style.display = 'block';
      }
      toast(`Table "${name}" applied (${rowConfigs.length} rows)`);
      closeTableModal();
      await loadSheet();
    } else {
      toast('Apply failed: ' + (data.error || 'unknown error'), true);
    }
  } catch(e) {
    toast('Error: ' + e.message, true);
  }
}

// Override cell click to handle shift+click for end_ref
const _origCellClick = window.cellClick;
window.cellClick = function(ref, e) {
  if (e && e.shiftKey && selectedAction === 'T') {
    handleShiftClick(ref);
    return;
  }
  if (_origCellClick) _origCellClick(ref, e);
};
