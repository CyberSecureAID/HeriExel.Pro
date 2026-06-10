/* ===============================================
   HeriExcel.Pro — app.js  v7.13 (CORREGIDO)
   Carga de archivos, renderizado de tabla,
   contexto, filtros, busqueda, zoom,
   exportacion, impresion, fijar filas,
   notas, fill-handle, redimensionado de columnas/filas

   FIX v7.13:
   - Se ha restaurado la función initColResize.
   - Se ha restaurado la función initRowResize.
   - Se ha asegurado que XLSXStyle se utilice para la lectura de archivos
     para que los estilos (colores, bordes, negritas, etc.) se carguen correctamente.
=============================================== */
'use strict';

/* ================================================
   INIT — único DOMContentLoaded
================================================ */
document.addEventListener('DOMContentLoaded', function() {
  initAutocomplete();
  initContextMenu();
  initFindReplace();
  initFilterDropdown();
  initRibbonActions();
  initMouseGlobalEvents();
  initTooltip();
  initFillHandle();
  initNoteDialog();
  renderCondRuleList();

  if (typeof initDatePicker   === 'function') initDatePicker();
  if (typeof initEyeDropper   === 'function') initEyeDropper();
  if (typeof initTemplateSelector === 'function') initTemplateSelector();

  initFileHandlers();
});

/* ================================================
   FILE HANDLERS
================================================ */
function initFileHandlers() {
  document.addEventListener('change', function(e) {
    if (e.target && e.target.id === 'fileInput') {
      var file = e.target.files[0];
      if (file) loadFile(file);
      e.target.value = '';
    }
  });

  document.addEventListener('dragover', function(e) {
    var dz = document.getElementById('dropZone');
    if (dz && dz.contains(e.target)) {
      e.preventDefault();
      dz.classList.add('drag-over');
    }
  });
  document.addEventListener('dragleave', function(e) {
    var dz = document.getElementById('dropZone');
    if (dz && !dz.contains(e.relatedTarget)) {
      dz.classList.remove('drag-over');
    }
  });
  document.addEventListener('drop', function(e) {
    var dz = document.getElementById('dropZone');
    if (dz && dz.contains(e.target)) {
      e.preventDefault();
      dz.classList.remove('drag-over');
      var file = e.dataTransfer.files[0];
      if (file) loadFile(file);
    }
  });

  var btnReset = document.getElementById('btnReset');
  if (btnReset) btnReset.addEventListener('click', resetApp);

  var btnRename = document.getElementById('btnRename');
  if (btnRename) btnRename.addEventListener('click', startRename);

  var fni = document.getElementById('fileNameInput');
  if (fni) {
    fni.addEventListener('blur', commitRename);
    fni.addEventListener('keydown', function(e) {
      if (e.key === 'Enter')  commitRename();
      if (e.key === 'Escape') cancelRename();
    });
  }
}

/* ================================================
   LOAD FILE — LECTURA MANUAL DE DATOS Y ESTILOS
================================================ */
function loadFile(file) {
  STATE.fileName = file.name;
  showLoading('Leyendo archivo...');

  var reader = new FileReader();
  reader.onload = function(e) {
    try {
      var data = new Uint8Array(e.target.result);
      var XL = window.XLSXStyle || XLSX;
      var wb = XL.read(data, { type: 'array', cellDates: true, cellStyles: true });

      STATE.workbook = wb;
      STATE.sheetNames = wb.SheetNames;
      loadSheet(wb.SheetNames[0]);
      showEditor();
    } catch (err) {
      hideLoading();
      showToast('Error al leer el archivo: ' + err.message, 'error');
      console.error('loadFile error:', err);
    }
  };
  reader.onerror = function() { hideLoading(); showToast('Error al leer el archivo', 'error'); };
  reader.readAsArrayBuffer(file);
}

function loadSheet(sheetName) {
  STATE.activeSheet = sheetName;
  var ws = STATE.workbook.Sheets[sheetName];
  STATE.wsRaw = ws;

  var XL = window.XLSXStyle || XLSX;
  var ref = ws['!ref'];
  if (!ref) {
    STATE.data = [['']];
    STATE.mergeMap = {};
  } else {
    var range = XL.utils.decode_range(ref);
    var numRows = range.e.r - range.s.r + 1;
    var numCols = range.e.c - range.s.c + 1;

    var data = [];
    for (var r = 0; r < numRows; r++) {
      data[r] = [];
      for (var c = 0; c < numCols; c++) {
        var addr = XL.utils.encode_cell({ r: r + range.s.r, c: c + range.s.c });
        var cell = ws[addr];
        var val = cell ? cell.v : '';
        data[r][c] = val;
      }
    }
    STATE.data = data;

    STATE.mergeMap = {};
    if (ws['!merges']) {
      ws['!merges'].forEach(function(m) {
        for (var mr = m.s.r; mr <= m.e.r; mr++) {
          for (var mc = m.s.c; mc <= m.e.c; mc++) {
            if (mr === m.s.r && mc === m.s.c) {
              STATE.mergeMap[mr + ',' + mc] = { rowspan: m.e.r - m.s.r + 1, colspan: m.e.c - m.s.c + 1 };
            } else {
              STATE.mergeMap[mr + ',' + mc] = 'skip';
            }
          }
        }
      });
    }
  }

  STATE.activeCellR = -1;
  STATE.activeCellC = -1;
  STATE.selectedCells.clear();
  STATE.selectedRow = -1;
  STATE.frozen = false;
  STATE.activeFilters = {};
  STATE.undoStack = [];
  STATE.redoStack = [];
  STATE.dirty = false;
  STATE.colWidths = {};
  STATE.rowHeights = {};

  _clearSheetFmt(sheetName);
  importStyles(ws, sheetName);

  if (ws['!cols']) {
    ws['!cols'].forEach(function(col, ci) {
      if (col && col.wch) STATE.colWidths[ci] = Math.round(col.wch * 7);
    });
  }

  renderSheetTabs();
  renderTable();
  updateStats();
  hideLoading();
  updateUnsaved();
}

/* ================================================
   LIMPIA FORMATOS DE UNA HOJA
================================================ */
function _clearSheetFmt(sheetName) {
  var prefix = sheetName + ':';
  Object.keys(FMT).forEach(function(k) {
    if (k.indexOf(prefix) === 0) delete FMT[k];
  });
  Object.keys(BORDERS).forEach(function(k) {
    if (k.indexOf(prefix) === 0) delete BORDERS[k];
  });
}

/* ================================================
   IMPORT STYLES — LEE TODOS LOS ESTILOS DE ws
================================================ */
function importStyles(ws, sheetName) {
  if (!ws || !sheetName) return;

  var numFmtInverse = {
    '#,##0.00': 'number',
    '#,##0': 'integer',
    '"$"#,##0.00': 'currency',
    '0.00%': 'percent',
    '0.00"%"': 'percent',
    '0.00E+00': 'scientific',
    'dd/mm/yyyy': 'date',
    'dd/mm/yy': 'date',
    'mm/dd/yy': 'date',
    'm/d/yy': 'date',
    'yyyy-mm-dd': 'date'
  };

  var XL = window.XLSXStyle || XLSX;
  var ref = ws['!ref'];
  if (!ref) return;
  var range = XL.utils.decode_range(ref);

  for (var r = range.s.r; r <= range.e.r; r++) {
    for (var c = range.s.c; c <= range.e.c; c++) {
      var addr = XL.utils.encode_cell({ r: r, c: c });
      var cell = ws[addr];
      if (!cell || !cell.s) continue;

      var s = cell.s;
      var fmt = {};
      var borders = { top: null, bottom: null, left: null, right: null };
      var hasFmt = false;
      var hasBorder = false;

      if (s.font) {
        if (s.font.bold) { fmt.bold = true; hasFmt = true; }
        if (s.font.italic) { fmt.italic = true; hasFmt = true; }
        if (s.font.underline) { fmt.underline = true; hasFmt = true; }
        if (s.font.strike) { fmt.strike = true; hasFmt = true; }
        if (s.font.name) { fmt.font = "'" + s.font.name + "', sans-serif"; hasFmt = true; }
        if (s.font.sz) { fmt.size = Math.round(s.font.sz / 0.75) + 'px'; hasFmt = true; }
        if (s.font.color && s.font.color.rgb) {
          fmt.color = '#' + s.font.color.rgb.slice(-6).toLowerCase();
          hasFmt = true;
        }
      }

      if (s.fill) {
        var rgb = null;
        if (s.fill.fgColor && s.fill.fgColor.rgb) rgb = s.fill.fgColor.rgb;
        else if (s.fill.bgColor && s.fill.bgColor.rgb) rgb = s.fill.bgColor.rgb;
        if (rgb && rgb !== '00000000' && rgb !== 'FFFFFFFF' && rgb !== 'FF000000') {
          fmt.bgColor = '#' + rgb.slice(-6).toLowerCase();
          hasFmt = true;
        }
      }

      if (s.alignment && s.alignment.horizontal) {
        var ha = s.alignment.horizontal;
        if (ha === 'left' || ha === 'center' || ha === 'right') {
          fmt.align = ha;
          hasFmt = true;
        }
      }

      if (s.numFmt) {
        var nf = numFmtInverse[s.numFmt];
        if (nf) { fmt.numFormat = nf; hasFmt = true; }
      }

      if (s.border) {
        var SIDES = ['top', 'bottom', 'left', 'right'];
        SIDES.forEach(function(side) {
          var bd = s.border[side];
          if (bd && bd.style) {
            var cssStyle = bd.style;
            if (cssStyle === 'medium' || cssStyle === 'thick') cssStyle = 'solid';
            if (cssStyle === 'thin') cssStyle = 'solid';
            var width = '1px';
            if (bd.style === 'medium') width = '2px';
            if (bd.style === 'thick') width = '3px';
            var color = '#7c8cf8';
            if (bd.color && bd.color.rgb) color = '#' + bd.color.rgb.slice(-6).toLowerCase();
            borders[side] = { color: color, style: cssStyle, width: width };
            hasBorder = true;
          }
        });
      }

      var key = sheetName + ':' + r + ',' + c;
      if (hasFmt) FMT[key] = fmt;
      if (hasBorder) BORDERS[key] = borders;
    }
  }
}

/* ================================================
   SHOW / HIDE EDITOR
================================================ */
function showEditor() {
  var uploadSection = document.getElementById('uploadSection');
  var editorSection = document.getElementById('editorSection');
  if (uploadSection) uploadSection.classList.add('hidden');
  if (editorSection) editorSection.classList.remove('hidden');
  document.getElementById('headerCenter').style.display = '';
  document.getElementById('btnReset').classList.remove('hidden');
  document.getElementById('btnExportMenu').classList.remove('hidden');
  document.getElementById('formatBar').classList.remove('hidden');
  document.getElementById('fxBar').classList.remove('hidden');
  document.getElementById('btnRename').style.display = '';
  document.getElementById('fileName').textContent = STATE.fileName;
  updateStats();
}

function resetApp() {
  STATE.workbook = null; STATE.data = []; STATE.sheetNames = [];
  STATE.activeSheet = ''; STATE.dirty = false;
  STATE.activeCellR = -1; STATE.activeCellC = -1;
  STATE.selectedCells.clear();
  STATE.undoStack = []; STATE.redoStack = [];
  STATE.colWidths = {}; STATE.rowHeights = {};

  var uploadSection = document.getElementById('uploadSection');
  var editorSection = document.getElementById('editorSection');
  if (uploadSection) uploadSection.classList.remove('hidden');
  if (editorSection) editorSection.classList.add('hidden');

  document.getElementById('headerCenter').style.display = 'none';
  document.getElementById('btnReset').classList.add('hidden');
  document.getElementById('btnExportMenu').classList.add('hidden');
  document.getElementById('formatBar').classList.add('hidden');
  document.getElementById('fxBar').classList.add('hidden');
  document.getElementById('tableHead').innerHTML = '';
  document.getElementById('tableBody').innerHTML = '';
  document.getElementById('sheetTabs').innerHTML = '';
  document.getElementById('fileName').textContent = '—';
  if (typeof dpHide === 'function') dpHide();
}

/* ================================================
   SHEET TABS
================================================ */
function renderSheetTabs() {
  var container = document.getElementById('sheetTabs');
  container.innerHTML = '';
  STATE.sheetNames.forEach(function(name) {
    var tab = document.createElement('button');
    tab.className = 'sheet-tab' + (name === STATE.activeSheet ? ' active' : '');
    tab.textContent = name;
    tab.addEventListener('click', function() {
      if (name !== STATE.activeSheet) {
        showLoading('Cargando hoja...');
        setTimeout(function() {
          try {
            loadSheet(name);
            renderSheetTabs();
          } catch (err) {
            hideLoading();
            showToast('Error al cargar hoja: ' + err.message, 'error');
            console.error('renderSheetTabs error:', err);
          }
        }, 30);
      }
    });
    container.appendChild(tab);
  });
}

/* ================================================
   RENDER TABLE (con initColResize y initRowResize)
================================================ */
function renderTable() {
  var thead = document.getElementById('tableHead');
  var tbody = document.getElementById('tableBody');
  thead.innerHTML = '';
  tbody.innerHTML = '';

  if (!STATE.data.length) return;

  var numCols = STATE.data[0] ? STATE.data[0].length : 0;

  var headRow = document.createElement('tr');
  var thCorner = document.createElement('th');
  thCorner.textContent = '#';
  headRow.appendChild(thCorner);

  for (var ci = 0; ci < numCols; ci++) {
    (function(c) {
      var th = document.createElement('th');
      th.setAttribute('data-col', c);

      var inner = document.createElement('div');
      inner.className = 'th-inner';

      var label = document.createElement('span');
      label.className = 'th-label';
      var headerVal = STATE.frozen
        ? ((STATE.data[0] ? STATE.data[0][c] : '') || columnLabel(c))
        : columnLabel(c);
      label.textContent = headerVal;
      label.title = headerVal;

      var filterIcon = document.createElement('span');
      filterIcon.className = 'filter-icon' + (STATE.activeFilters[c] ? ' active' : '');
      filterIcon.innerHTML = '<svg viewBox="0 0 20 20" fill="currentColor"><path fill-rule="evenodd" d="M3 3a1 1 0 011-1h12a1 1 0 011 1v3a1 1 0 01-.293.707L13 10.414V17a1 1 0 01-1.447.894l-4-2A1 1 0 017 15v-4.586L3.293 6.707A1 1 0 013 6V3z" clip-rule="evenodd"/></svg>';
      filterIcon.addEventListener('click', function(e) {
        e.stopPropagation();
        openFilterDropdown(c, filterIcon);
      });

      var resize = document.createElement('div');
      resize.className = 'resize-handle';
      initColResize(th, c, resize);

      inner.appendChild(label);
      inner.appendChild(filterIcon);
      th.appendChild(inner);
      th.appendChild(resize);

      th.addEventListener('click', function(e) {
        if (!e.target.closest('.resize-handle') && !e.target.closest('.filter-icon')) {
          selectEntireCol(c);
        }
      });

      if (STATE.colWidths[c]) {
        th.style.minWidth = STATE.colWidths[c] + 'px';
        th.style.width = STATE.colWidths[c] + 'px';
      }

      headRow.appendChild(th);
    })(ci);
  }
  thead.appendChild(headRow);

  var startRow = STATE.frozen ? 1 : 0;
  var frag = document.createDocumentFragment();

  for (var ri = startRow; ri < STATE.data.length; ri++) {
    (function(r) {
      var tr = document.createElement('tr');
      tr.setAttribute('data-row', r);
      if (isRowHidden(r)) tr.classList.add('row-hidden');
      if (r === STATE.selectedRow) tr.classList.add('selected-row');

      var rowNumTd = document.createElement('td');
      rowNumTd.className = 'row-num';
      rowNumTd.textContent = r + 1;
      rowNumTd.setAttribute('data-row', r);

      var rowResizeHandle = document.createElement('div');
      rowResizeHandle.className = 'row-resize-handle';
      initRowResize(tr, r, rowResizeHandle);
      rowNumTd.appendChild(rowResizeHandle);
      rowNumTd.addEventListener('click', function() { selectEntireRow(r); });
      tr.appendChild(rowNumTd);

      for (var cj = 0; cj < numCols; cj++) {
        (function(c) {
          var mergeInfo = STATE.mergeMap[r + ',' + c];
          if (mergeInfo === 'skip') return;

          var td = document.createElement('td');
          td.setAttribute('data-col', c);
          td.setAttribute('data-row', r);

          if (mergeInfo && typeof mergeInfo === 'object') {
            td.rowSpan = mergeInfo.rowspan;
            td.colSpan = mergeInfo.colspan;
            td.classList.add('merged-cell');
          }

          var rawVal = (STATE.data[r] ? STATE.data[r][c] : '') || '';
          var fmt = getCellFmt(r, c);

          var span = document.createElement('span');
          span.className = 'cell-display';

          var displayVal = rawVal;
          var isFormula = typeof rawVal === 'string' && rawVal.startsWith('=');
          if (isFormula) {
            displayVal = evaluateFormula(rawVal, r, c);
            var isErr = displayVal === '#ERROR!' || displayVal === '#DIV/0!' || displayVal === '#N/A' || displayVal === '#NUM!';
            if (isErr) span.classList.add('formula-error');
          }
          span.textContent = formatCellValue(displayVal, fmt.numFormat);
          applyFmtToSpan(span, fmt);
          applyCondFmtToCell(span, r, c, rawVal);
          applyBordersToTd(td, r, c);

          if (STATE.selectedCells.has(r + ',' + c)) {
            td.classList.add('td-selected');
            if (r === STATE.activeCellR && c === STATE.activeCellC) {
              td.classList.add('td-active');
            }
          }

          if (STATE.colWidths[c]) {
            td.style.width = STATE.colWidths[c] + 'px';
            td.style.minWidth = STATE.colWidths[c] + 'px';
            td.style.maxWidth = STATE.colWidths[c] + 'px';
          }
          if (STATE.rowHeights[r]) tr.style.height = STATE.rowHeights[r] + 'px';

          td.appendChild(span);
          renderNoteIndicator(r, c, td);

          td.addEventListener('mousedown', function(e) {
            var rect2 = td.getBoundingClientRect();
            var isHandle = (e.clientX >= rect2.right - 8) && (e.clientY >= rect2.bottom - 8);
            if (isHandle && e.button === 0 &&
                STATE.activeCellR === r && STATE.activeCellC === c &&
                !STATE.editingCell) {
              e.preventDefault();
              e.stopPropagation();
              STATE.isFillDragging = true;
              STATE.fillStart = { r: r, c: c };
              STATE.fillEnd = { r: r, c: c };
              STATE.fillAnchorR = r;
              STATE.fillAnchorC = c;
              return;
            }
            handleCellMousedown(e, r, c);
          });
          td.addEventListener('mouseenter', function(e) {
            if (STATE.isFillDragging) {
              STATE.fillEnd = { r: r, c: c };
              highlightFillRange(STATE.fillStart, { r: r, c: c });
              return;
            }
            handleCellMouseenter(e, r, c);
          });
          td.addEventListener('dblclick', function() { enterEdit(r, c); });
          td.addEventListener('contextmenu', function(e) { e.preventDefault(); showContextMenu(e, r, c); });

          tr.appendChild(td);
        })(cj);
      }
      frag.appendChild(tr);
    })(ri);
  }
  tbody.appendChild(frag);

  applyZoom();
  updateColHighlight();
  updateFreezeLine();
}

/* ────────────────────────────────────────────
   INIT COL RESIZE
───────────────────────────────────────────── */
function initColResize(th, colIdx, handle) {
  handle.addEventListener('mousedown', function(e) {
    e.preventDefault(); e.stopPropagation();
    var startX = e.clientX;
    var startW = STATE.colWidths[colIdx] !== undefined ? STATE.colWidths[colIdx] : (th.offsetWidth || 90);
    document.body.style.cursor = 'col-resize';
    document.body.style.userSelect = 'none';

    var onMove = function(ev) {
      var w = Math.max(20, startW + ev.clientX - startX);
      STATE.colWidths[colIdx] = w;
      th.style.minWidth = w + 'px';
      th.style.width = w + 'px';
      th.style.maxWidth = w + 'px';
      document.querySelectorAll('td[data-col="' + colIdx + '"]').forEach(function(td) {
        td.style.width = w + 'px';
        td.style.minWidth = w + 'px';
        td.style.maxWidth = w + 'px';
      });
    };
    var onUp = function() {
      document.body.style.cursor = '';
      document.body.style.userSelect = '';
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
    };
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
  });
}

/* ────────────────────────────────────────────
   INIT ROW RESIZE
───────────────────────────────────────────── */
function initRowResize(tr, rowIdx, handle) {
  handle.addEventListener('mousedown', function(e) {
    e.preventDefault(); e.stopPropagation();
    var startY = e.clientY;
    var startH = STATE.rowHeights[rowIdx] || tr.offsetHeight || 28;
    document.body.style.cursor = 'row-resize';

    var onMove = function(ev) {
      var h = Math.max(18, startH + ev.clientY - startY);
      tr.style.height = h + 'px';
      STATE.rowHeights[rowIdx] = h;
    };
    var onUp = function() {
      document.body.style.cursor = '';
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
    };
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
  });
}

/* ================================================
   FUNCIONES AUXILIARES
================================================ */
function isRowHidden(r) {
  if (!Object.keys(STATE.activeFilters).length) return false;
  for (var colStr in STATE.activeFilters) {
    var c = parseInt(colStr);
    var val = String((STATE.data[r] ? STATE.data[r][c] : '') || '');
    if (STATE.activeFilters[colStr].indexOf(val) === -1) return true;
  }
  return false;
}

function highlightFillRange(start, end) {
  if (!start || !end) return;
  var tbody = document.getElementById('tableBody');
  if (!tbody) return;
  tbody.querySelectorAll('.td-fill-preview').forEach(function(el) {
    el.classList.remove('td-fill-preview');
    el.style.outline = '';
  });
  var minR = Math.min(start.r, end.r), maxR = Math.max(start.r, end.r);
  for (var r = minR; r <= maxR; r++) {
    var td = getCellEl(r, start.c);
    if (td && r !== start.r) {
      td.classList.add('td-fill-preview');
      td.style.outline = '1px dashed var(--accent)';
    }
  }
}

function updateFreezeLine() {
  var line = document.querySelector('.freeze-col-line');
  if (!STATE.frozen) { if (line) line.remove(); return; }
}

function updateStats() {
  var rows = STATE.data.length;
  var cols = STATE.data[0] ? STATE.data[0].length : 0;
  document.getElementById('statRows').textContent = rows + ' filas';
  document.getElementById('statCols').textContent = cols + ' columnas';
}

function markDirty() {
  STATE.dirty = true;
  updateUnsaved();
}

function updateUnsaved() {
  var dot = document.getElementById('fileDot');
  var sep = document.getElementById('unsavedSep');
  var stat = document.getElementById('statUnsaved');
  var show = STATE.dirty;
  if (dot) dot.classList.toggle('visible', show);
  if (sep) sep.style.display = show ? '' : 'none';
  if (stat) stat.style.display = show ? '' : 'none';
}

function applyZoom() {
  var table = document.getElementById('excelTable');
  var scroll = document.getElementById('tableScroll');
  if (!table || !scroll) return;

  var scale = STATE.zoom / 100;
  table.style.transform = 'scale(1)';
  table.style.transformOrigin = 'top left';

  var naturalW = table.scrollWidth;
  var naturalH = table.scrollHeight;

  table.style.transform = 'scale(' + scale + ')';
  scroll.style.paddingRight = Math.max(0, naturalW * scale - scroll.clientWidth + 20) + 'px';
  scroll.style.paddingBottom = Math.max(0, naturalH * scale - scroll.clientHeight + 20) + 'px';
  document.getElementById('zoomIndicator').textContent = STATE.zoom + '%';
}

/* ================================================
   EXPORT — XLSX con estilos completos
================================================ */
function _cssColorToArgb(css) {
  if (!css) return null;
  css = css.trim();
  if (/^#[0-9a-f]{3}$/i.test(css)) {
    css = '#' + css[1] + css[1] + css[2] + css[2] + css[3] + css[3];
  }
  if (/^#[0-9a-f]{6}$/i.test(css)) {
    return 'FF' + css.slice(1).toUpperCase();
  }
  var m = css.match(/^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i);
  if (m) {
    return 'FF' +
      parseInt(m[1]).toString(16).padStart(2, '0').toUpperCase() +
      parseInt(m[2]).toString(16).padStart(2, '0').toUpperCase() +
      parseInt(m[3]).toString(16).padStart(2, '0').toUpperCase();
  }
  return null;
}

function _resolveCssVar(val) {
  if (!val) return val;
  val = val.trim();
  if (!val.startsWith('var(')) return val;
  var m = val.match(/^var\(\s*(--[^,)]+)/);
  if (!m) return val;
  var resolved = getComputedStyle(document.documentElement).getPropertyValue(m[1]).trim();
  return resolved || val;
}

function _buildCellStyle(r, c) {
  var fmtKey = STATE.activeSheet + ':' + r + ',' + c;
  var fmt = FMT[fmtKey] || {};
  var borders = BORDERS[fmtKey] || {};

  var hasFmt = fmt.bold || fmt.italic || fmt.underline || fmt.strike ||
               fmt.font || fmt.size || fmt.color || fmt.bgColor ||
               fmt.align || (fmt.numFormat && fmt.numFormat !== 'general') ||
               fmt.cellStyle;

  var hasBorder = (borders.top && borders.top.color) ||
                  (borders.bottom && borders.bottom.color) ||
                  (borders.left && borders.left.color) ||
                  (borders.right && borders.right.color);

  if (!hasFmt && !hasBorder) return null;

  var style = {};

  var fontObj = {};
  if (fmt.bold) fontObj.bold = true;
  if (fmt.italic) fontObj.italic = true;
  if (fmt.underline) fontObj.underline = true;
  if (fmt.strike) fontObj.strike = true;
  if (fmt.font) {
    var fontName = fmt.font.replace(/'/g, '').split(',')[0].trim();
    if (fontName) fontObj.name = fontName;
  }
  if (fmt.size) {
    var pxVal = parseFloat(fmt.size);
    if (!isNaN(pxVal)) fontObj.sz = Math.round(pxVal * 0.75);
  }
  if (fmt.color) {
    var fc = _cssColorToArgb(_resolveCssVar(fmt.color));
    if (fc) fontObj.color = { rgb: fc };
  }
  if (Object.keys(fontObj).length) style.font = fontObj;

  var bgColor = fmt.bgColor;
  if (!bgColor && fmt.cellStyle) {
    var styleColorMap = {
      header: '#2d3050', total: '#2d3050', good: '#1a3a1a',
      bad: '#3a1a1a', neutral: '#3a350a', warning: '#3a2a0a'
    };
    bgColor = styleColorMap[fmt.cellStyle] || null;
  }
  if (bgColor) {
    var bc = _cssColorToArgb(_resolveCssVar(bgColor));
    if (bc) style.fill = { patternType: 'solid', fgColor: { rgb: bc } };
  }

  if (fmt.align && (fmt.align === 'left' || fmt.align === 'center' || fmt.align === 'right')) {
    style.alignment = { horizontal: fmt.align };
  }

  var numFmtMap = {
    number: '#,##0.00', integer: '#,##0', currency: '"$"#,##0.00',
    percent: '0.00"%"', scientific: '0.00E+00', date: 'dd/mm/yyyy'
  };
  if (fmt.numFormat && numFmtMap[fmt.numFormat]) {
    style.numFmt = numFmtMap[fmt.numFormat];
  }

  if (hasBorder) {
    var borderObj = {};
    ['top', 'bottom', 'left', 'right'].forEach(function(side) {
      var bd = borders[side];
      if (!bd || !bd.color) return;
      var xlStyle = 'thin';
      if (bd.style === 'dashed') xlStyle = 'dashed';
      if (bd.style === 'double') xlStyle = 'double';
      var wNum = parseFloat(bd.width) || 1;
      if (wNum >= 3) xlStyle = 'thick';
      else if (wNum >= 2) xlStyle = 'medium';
      var bColor = _cssColorToArgb(_resolveCssVar(bd.color));
      var entry = { style: xlStyle };
      if (bColor) entry.color = { rgb: bColor };
      borderObj[side] = entry;
    });
    if (Object.keys(borderObj).length) style.border = borderObj;
  }

  return Object.keys(style).length ? style : null;
}

function exportXLSX() {
  if (!STATE.data.length) return showToast('No hay datos para exportar', '');

  var XL = window.XLSXStyle || XLSX;
  var hasStyles = !!(window.XLSXStyle);

  showLoading('Exportando...');
  setTimeout(function() {
    try {
      var wb = XL.utils.book_new();
      var ws = XL.utils.aoa_to_sheet(STATE.data);

      if (hasStyles) {
        var numRows = STATE.data.length;
        var numCols = STATE.data[0] ? STATE.data[0].length : 0;
        for (var r = 0; r < numRows; r++) {
          for (var c = 0; c < numCols; c++) {
            var cellStyle = _buildCellStyle(r, c);
            if (!cellStyle) continue;
            var cellAddr = XL.utils.encode_cell({ r: r, c: c });
            if (!ws[cellAddr]) ws[cellAddr] = { t: 's', v: '' };
            ws[cellAddr].s = cellStyle;
          }
        }
      }

      var numCols2 = STATE.data[0] ? STATE.data[0].length : 0;
      var colInfo = [];
      for (var ci = 0; ci < numCols2; ci++) {
        var w = STATE.colWidths[ci];
        colInfo.push(w ? { wch: Math.round(w / 7) } : { wch: 14 });
      }
      ws['!cols'] = colInfo;

      XL.utils.book_append_sheet(wb, ws, STATE.activeSheet || 'Hoja1');

      var fname = (STATE.fileName.replace(/\.[^.]+$/, '') || 'documento') + '.xlsx';

      if (hasStyles) {
        XL.writeFile(wb, fname, { bookSST: true, cellStyles: true });
        showToast('Archivo guardado con estilos: ' + fname, 'success');
      } else {
        XL.writeFile(wb, fname);
        showToast('Archivo guardado: ' + fname, 'info');
      }

      STATE.dirty = false;
      updateUnsaved();
    } catch (err) {
      showToast('Error al exportar: ' + err.message, 'error');
      console.error('exportXLSX error:', err);
    }
    hideLoading();
  }, 50);
}

function exportCSV() {
  if (!STATE.data.length) return showToast('No hay datos para exportar', '');
  var csv = STATE.data.map(function(row) {
    return row.map(function(v) {
      var s = String(v);
      return (s.indexOf(',') !== -1 || s.indexOf('"') !== -1 || s.indexOf('\n') !== -1)
        ? '"' + s.replace(/"/g, '""') + '"' : s;
    }).join(',');
  }).join('\n');
  var blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' });
  var url = URL.createObjectURL(blob);
  var a = document.createElement('a');
  a.href = url;
  a.download = (STATE.fileName.replace(/\.[^.]+$/, '') || 'datos') + '.csv';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  showToast('CSV exportado', 'success');
}

/* ================================================
   SEARCH
================================================ */
var _searchMatches = [], _searchIdx = -1;

function performSearch(q) {
  _searchMatches = []; _searchIdx = -1;
  var lower = q.toLowerCase();
  var tbody = document.getElementById('tableBody');
  if (!tbody) return;

  tbody.querySelectorAll('.cell-display').forEach(function(span) {
    var td = span.closest('td'); if (!td) return;
    var r = parseInt(td.getAttribute('data-row'));
    var c = parseInt(td.getAttribute('data-col'));
    if (isNaN(r) || isNaN(c)) return;
    var rawSearch = String((STATE.data[r] ? STATE.data[r][c] : '') || '');
    if (rawSearch.toLowerCase().indexOf(lower) !== -1) {
      span.classList.add('search-match');
      _searchMatches.push({ span: span, r: r, c: c });
    }
  });

  document.getElementById('searchCountBadge').textContent = _searchMatches.length ? String(_searchMatches.length) : '0';
  if (_searchMatches.length) highlightSearchResult(0);
}

function clearSearchHighlights() {
  document.querySelectorAll('.search-match, .search-current').forEach(function(el) {
    el.classList.remove('search-match', 'search-current');
  });
  _searchMatches = []; _searchIdx = -1;
}

function highlightSearchResult(idx) {
  _searchMatches.forEach(function(m) { m.span.classList.remove('search-current'); });
  if (!_searchMatches.length) return;
  _searchIdx = ((idx % _searchMatches.length) + _searchMatches.length) % _searchMatches.length;
  var cur = _searchMatches[_searchIdx];
  cur.span.classList.add('search-current');
  cur.span.scrollIntoView({ block: 'nearest', inline: 'nearest' });
}

/* ================================================
   FIND & REPLACE PANEL
================================================ */
var _frpMatches = [], _frpIdx = -1;

function showFindReplace() { document.getElementById('findReplacePanel').classList.remove('hidden'); document.getElementById('frpFind').focus(); }
function hideFindReplace() { document.getElementById('findReplacePanel').classList.add('hidden'); }

function frpSearch() {
  var q = document.getElementById('frpFind').value;
  var caseSen = document.getElementById('frpCaseSensitive').checked;
  var whole = document.getElementById('frpWholeCell').checked;
  _frpMatches = []; _frpIdx = -1;
  clearSearchHighlights();
  if (!q) { document.getElementById('frpInfo').textContent = ''; document.getElementById('frpInfo').className = 'frp-info'; return; }

  STATE.data.forEach(function(row, r) {
    row.forEach(function(val, c) {
      var v = String(val), s = q;
      if (!caseSen) { v = v.toLowerCase(); s = s.toLowerCase(); }
      var match = whole ? v === s : v.indexOf(s) !== -1;
      if (match) _frpMatches.push({ r: r, c: c });
    });
  });

  var info = document.getElementById('frpInfo');
  if (_frpMatches.length) {
    info.textContent = _frpMatches.length + ' coincidencia(s)';
    info.className = 'frp-info found';
    frpGoTo(0);
  } else {
    info.textContent = 'No encontrado';
    info.className = 'frp-info none';
  }
}

function frpGoTo(idx) {
  if (!_frpMatches.length) return;
  _frpIdx = ((idx % _frpMatches.length) + _frpMatches.length) % _frpMatches.length;
  var match = _frpMatches[_frpIdx];
  activateCell(match.r, match.c);
  var span2 = getCellSpan(match.r, match.c);
  if (span2) { clearSearchHighlights(); span2.classList.add('search-current'); }
  document.getElementById('frpInfo').textContent = (_frpIdx + 1) + ' / ' + _frpMatches.length;
}

function initFindReplace() {
  var frpClose = document.getElementById('frpClose');
  var frpFind = document.getElementById('frpFind');
  var frpFindNext = document.getElementById('frpFindNext');
  var frpFindPrev = document.getElementById('frpFindPrev');
  var frpCaseSensitive = document.getElementById('frpCaseSensitive');
  var frpWholeCell = document.getElementById('frpWholeCell');
  var frpReplaceOne = document.getElementById('frpReplaceOne');
  var frpReplaceAll = document.getElementById('frpReplaceAll');

  if (frpClose) frpClose.addEventListener('click', hideFindReplace);
  if (frpFind) frpFind.addEventListener('input', frpSearch);
  if (frpFindNext) frpFindNext.addEventListener('click', function() { frpGoTo(_frpIdx + 1); });
  if (frpFindPrev) frpFindPrev.addEventListener('click', function() { frpGoTo(_frpIdx - 1); });
  if (frpCaseSensitive) frpCaseSensitive.addEventListener('change', frpSearch);
  if (frpWholeCell) frpWholeCell.addEventListener('change', frpSearch);

  if (frpReplaceOne) {
    frpReplaceOne.addEventListener('click', function() {
      if (_frpIdx < 0 || !_frpMatches.length) return frpSearch();
      var match2 = _frpMatches[_frpIdx];
      var q2 = document.getElementById('frpFind').value;
      var rep2 = document.getElementById('frpReplace').value;
      var caseSen2 = document.getElementById('frpCaseSensitive').checked;
      var whole2 = document.getElementById('frpWholeCell').checked;
      pushUndo();
      if (whole2) {
        STATE.data[match2.r][match2.c] = rep2;
      } else {
        var flags2 = caseSen2 ? 'g' : 'gi';
        var esc2 = q2.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        STATE.data[match2.r][match2.c] = String(STATE.data[match2.r][match2.c])
          .replace(new RegExp(esc2, flags2), rep2);
      }
      renderTable(); markDirty(); frpSearch();
      showToast('Reemplazado', 'success');
    });
  }
  if (frpReplaceAll) {
    frpReplaceAll.addEventListener('click', function() {
      var q3 = document.getElementById('frpFind').value;
      var rep3 = document.getElementById('frpReplace').value;
      var caseSen3 = document.getElementById('frpCaseSensitive').checked;
      var whole3 = document.getElementById('frpWholeCell').checked;
      if (!q3) return showToast('Escribe un término a buscar', '');
      pushUndo();
      var count2 = 0;
      var flags3 = caseSen3 ? 'g' : 'gi';
      var esc3 = q3.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      var regex3 = new RegExp(esc3, flags3);
      STATE.data.forEach(function(row, r) {
        row.forEach(function(val, c) {
          var v = String(val), s = q3;
          if (!caseSen3) { v = v.toLowerCase(); s = s.toLowerCase(); }
          var matches3 = whole3 ? v === s : v.indexOf(s) !== -1;
          if (matches3) {
            STATE.data[r][c] = whole3 ? rep3 : String(STATE.data[r][c]).replace(regex3, rep3);
            count2++;
          }
        });
      });
      renderTable(); markDirty(); frpSearch();
      showToast(count2 + ' reemplazo(s)', 'success');
    });
  }
}

/* ================================================
   CONTEXT MENU
================================================ */
function initContextMenu() {
  document.addEventListener('click', function(e) {
    if (!e.target.closest('#contextMenu')) hideContextMenu();
  });
}

function showContextMenu(e, r, c) {
  activateCell(r, c);
  var ctx = document.getElementById('contextMenu');
  ctx.innerHTML = '';

  var menuItems = [
    { label: 'Cortar', icon: '&#9986;', action: function() { copyCells(true); } },
    { label: 'Copiar', icon: '&#128203;', action: function() { copyCells(false); } },
    { label: 'Pegar', icon: '&#128204;', action: function() { pasteCells(false); } },
    { label: 'Pegar solo valores', icon: '&#128204;', action: function() { pasteCells(true); } },
    { sep: true },
    {
      label: 'Insertar fila arriba', icon: '&#43;',
      action: function() {
        pushUndo();
        var cols = STATE.data[0] ? STATE.data[0].length : 1;
        shiftRowKeys(r, +1, cols);
        STATE.data.splice(r, 0, new Array(cols).fill(''));
        STATE.activeCellR = r + 1;
        renderTable(); updateStats(); markDirty();
        showToast('Fila insertada arriba', 'success');
      }
    },
    {
      label: 'Insertar fila abajo', icon: '&#43;',
      action: function() {
        pushUndo();
        var cols2 = STATE.data[0] ? STATE.data[0].length : 1;
        shiftRowKeys(r + 1, +1, cols2);
        STATE.data.splice(r + 1, 0, new Array(cols2).fill(''));
        renderTable(); updateStats(); markDirty();
        showToast('Fila insertada abajo', 'success');
      }
    },
    {
      label: 'Eliminar fila', icon: '&#128465;', cls: 'danger',
      action: function() {
        if (STATE.data.length <= 1) return showToast('No se puede eliminar la única fila', '');
        pushUndo();
        var cols3 = STATE.data[0] ? STATE.data[0].length : 1;
        shiftRowKeys(r, -1, cols3);
        STATE.data.splice(r, 1);
        STATE.activeCellR = Math.min(r, STATE.data.length - 1);
        STATE.selectedCells.clear();
        renderTable(); updateStats(); markDirty();
        showToast('Fila eliminada', 'info');
      }
    },
    { sep: true },
    {
      label: 'Insertar columna antes', icon: '&#43;',
      action: function() {
        pushUndo();
        shiftColKeys(c, +1);
        STATE.data.forEach(function(row) { row.splice(c, 0, ''); });
        var nw = {};
        Object.keys(STATE.colWidths).forEach(function(k) {
          var ki = parseInt(k);
          if (ki >= c) nw[ki + 1] = STATE.colWidths[k];
          else nw[ki] = STATE.colWidths[k];
        });
        STATE.colWidths = nw;
        renderTable(); updateStats(); markDirty();
        showToast('Columna insertada antes', 'success');
      }
    },
    {
      label: 'Insertar columna después', icon: '&#43;',
      action: function() {
        pushUndo();
        shiftColKeys(c + 1, +1);
        STATE.data.forEach(function(row) { row.splice(c + 1, 0, ''); });
        var nw2 = {};
        Object.keys(STATE.colWidths).forEach(function(k) {
          var ki = parseInt(k);
          if (ki > c) nw2[ki + 1] = STATE.colWidths[k];
          else nw2[ki] = STATE.colWidths[k];
        });
        STATE.colWidths = nw2;
        renderTable(); updateStats(); markDirty();
        showToast('Columna insertada después', 'success');
      }
    },
    {
      label: 'Eliminar columna', icon: '&#128465;', cls: 'danger',
      action: function() {
        var numC = STATE.data[0] ? STATE.data[0].length : 0;
        if (numC <= 1) return showToast('No se puede eliminar la única columna', '');
        pushUndo();
        shiftColKeys(c, -1);
        STATE.data.forEach(function(row) { row.splice(c, 1); });
        var nw3 = {};
        Object.keys(STATE.colWidths).forEach(function(k) {
          var ki = parseInt(k);
          if (ki < c) nw3[ki] = STATE.colWidths[k];
          else if (ki > c) nw3[ki - 1] = STATE.colWidths[k];
        });
        STATE.colWidths = nw3;
        STATE.activeCellC = Math.max(0, c - 1);
        STATE.selectedCells.clear();
        renderTable(); updateStats(); markDirty();
        showToast('Columna eliminada', 'info');
      }
    },
    { sep: true },
    { label: 'Agregar nota', icon: '&#128221;', action: function() { addNote(r, c); } },
    { label: 'Limpiar celdas', icon: '&#10006;', cls: 'danger', action: function() { clearActiveCells(); } },
  ];

  menuItems.forEach(function(item) {
    if (item.sep) {
      var sep = document.createElement('div');
      sep.className = 'ctx-sep';
      ctx.appendChild(sep);
      return;
    }
    var btn = document.createElement('button');
    btn.className = 'ctx-item' + (item.cls ? ' ' + item.cls : '');
    btn.innerHTML = '<span style="min-width:16px;text-align:center;font-size:12px">' + (item.icon || '') + '</span> ' + escapeHtml(item.label);
    btn.addEventListener('click', function(e3) {
      e3.stopPropagation();
      hideContextMenu();
      setTimeout(function() { item.action(); }, 0);
    });
    ctx.appendChild(btn);
  });

  ctx.classList.remove('hidden');
  requestAnimationFrame(function() {
    var cw = ctx.offsetWidth || 200;
    var ch = ctx.offsetHeight || 300;
    var x = Math.min(e.clientX, window.innerWidth - cw - 10);
    var y = Math.min(e.clientY, window.innerHeight - ch - 10);
    ctx.style.left = Math.max(4, x) + 'px';
    ctx.style.top = Math.max(4, y) + 'px';
  });
}

function hideContextMenu() {
  var ctx = document.getElementById('contextMenu');
  if (ctx) ctx.classList.add('hidden');
}

/* ================================================
   FILTER DROPDOWN
================================================ */
var _filterCol = -1;

function initFilterDropdown() {
  document.getElementById('filterApply').addEventListener('click', applyColumnFilter);
  document.getElementById('filterClear').addEventListener('click', function() {
    delete STATE.activeFilters[_filterCol];
    hideFilterDropdown();
    renderTable();
    var th = document.querySelector('th[data-col="' + _filterCol + '"]');
    if (th) {
      var fi = th.querySelector('.filter-icon');
      if (fi) fi.classList.remove('active');
    }
    showToast('Filtro eliminado', 'info');
  });
  document.getElementById('filterSearch').addEventListener('input', function() {
    var q = document.getElementById('filterSearch').value.toLowerCase();
    document.getElementById('filterOptions').querySelectorAll('.filter-option').forEach(function(opt) {
      var txt = opt.querySelector('.filter-option-text');
      opt.style.display = txt && txt.textContent.toLowerCase().indexOf(q) !== -1 ? '' : 'none';
    });
  });
}

function openFilterDropdown(c, iconEl) {
  _filterCol = c;
  var dd = document.getElementById('filterDropdown');
  var vals = [];
  var seen = {};
  STATE.data.forEach(function(row) {
    var v = String(row[c] || '');
    if (!seen[v]) { seen[v] = true; vals.push(v); }
  });
  vals.sort();
  var cur = STATE.activeFilters[c] || vals;

  document.getElementById('filterSearch').value = '';
  var opts = document.getElementById('filterOptions');
  opts.innerHTML = '';
  vals.forEach(function(v) {
    var div = document.createElement('div');
    div.className = 'filter-option';
    var cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.value = v;
    cb.checked = cur.indexOf(v) !== -1;
    var txt = document.createElement('span');
    txt.className = 'filter-option-text';
    txt.textContent = v || '(vacío)';
    div.appendChild(cb);
    div.appendChild(txt);
    opts.appendChild(div);
  });

  var rect = iconEl.getBoundingClientRect();
  dd.style.left = Math.min(rect.left, window.innerWidth - 290) + 'px';
  dd.style.top = (rect.bottom + 4) + 'px';
  dd.classList.remove('hidden');
}

function applyColumnFilter() {
  var checkboxes = document.getElementById('filterOptions').querySelectorAll('input[type="checkbox"]:checked');
  var vals = Array.from(checkboxes).map(function(cb) { return cb.value; });
  if (!vals.length) { showToast('Selecciona al menos un valor para filtrar', ''); return; }
  STATE.activeFilters[_filterCol] = vals;
  hideFilterDropdown();
  renderTable();
  showToast('Filtro aplicado', 'success');
}

function hideFilterDropdown() { document.getElementById('filterDropdown').classList.add('hidden'); }

/* ================================================
   FILL HANDLE (drag to fill)
================================================ */
function initFillHandle() {
  var scroll = document.getElementById('tableScroll');
  scroll.addEventListener('mouseup', function() {
    if (STATE.isFillDragging && STATE.fillStart && STATE.fillEnd) {
      executeFillDown();
    }
    var tbody = document.getElementById('tableBody');
    if (tbody) {
      tbody.querySelectorAll('.td-fill-preview').forEach(function(el) {
        el.classList.remove('td-fill-preview');
        el.style.outline = '';
      });
    }
    STATE.isFillDragging = false;
    STATE.fillStart = null;
    STATE.fillEnd = null;
    STATE.isDragging = false;
  });
}

function executeFillDown() {
  var r0 = STATE.fillStart.r, c0 = STATE.fillStart.c;
  var r1 = STATE.fillEnd.r;
  if (r1 === r0) return;

  var direction = r1 > r0 ? 1 : -1;
  var minR = Math.min(r0, r1), maxR = Math.max(r0, r1);
  pushUndo();

  var seedVals = [];
  if (r0 > 0 && direction > 0) seedVals.push((STATE.data[r0 - 1] ? STATE.data[r0 - 1][c0] : '') || '');
  seedVals.push((STATE.data[r0] ? STATE.data[r0][c0] : '') || '');
  seedVals = seedVals.filter(function(v) { return v !== ''; });

  var seq = detectSequence(seedVals);
  for (var r = minR + 1; r <= maxR; r++) {
    if (seq) {
      STATE.data[r][c0] = getNextSequenceValue(seq, r - r0);
    } else {
      STATE.data[r][c0] = (STATE.data[r0] ? STATE.data[r0][c0] : '') || '';
    }
    var span = getCellSpan(r, c0);
    if (span) span.textContent = formatCellValue(STATE.data[r][c0], getCellFmt(r, c0).numFormat);
  }
  markDirty();
  showToast('Relleno aplicado', 'success');
}

/* ================================================
   GLOBAL MOUSE EVENTS
================================================ */
function initMouseGlobalEvents() {
  document.addEventListener('mouseup', function() {
    STATE.isDragging = false;
    STATE.isFillDragging = false;
  });
  document.addEventListener('mouseleave', function() {
    STATE.isDragging = false;
    STATE.isFillDragging = false;
  });
}

/* ================================================
   TOOLTIP
================================================ */
function initTooltip() {
  var tooltip = document.getElementById('cellTooltip');
  if (!tooltip) return;
  var _ttTimer;

  document.addEventListener('mouseover', function(e) {
    var td = e.target.closest('#tableBody td[data-col]');
    if (!td) return;
    var r = parseInt(td.getAttribute('data-row'));
    var c = parseInt(td.getAttribute('data-col'));
    if (isNaN(r) || isNaN(c) || r < 0 || c < 0) return;
    var raw = (STATE.data[r] ? STATE.data[r][c] : '') || '';
    var note = (STATE.cellNotes ? STATE.cellNotes[r + ',' + c] : null);
    if (!raw && !note) return;
    clearTimeout(_ttTimer);
    _ttTimer = setTimeout(function() {
      var html = '';
      if (note) html += '<div class="tooltip-label">Nota</div>' + escapeHtml(note) + '<br>';
      var rawStr = String(raw);
      if (rawStr && rawStr.startsWith('=')) html += '<div class="tooltip-label">Fórmula</div>' + escapeHtml(rawStr);
      else if (rawStr.length > 30) html += escapeHtml(rawStr);
      if (!html) return;
      tooltip.innerHTML = html;
      tooltip.style.left = Math.min(e.clientX + 12, window.innerWidth - 250) + 'px';
      tooltip.style.top = Math.min(e.clientY + 18, window.innerHeight - 90) + 'px';
      tooltip.classList.add('show');
    }, 600);
  });

  document.addEventListener('mouseout', function(e) {
    if (!e.target.closest('#tableBody')) return;
    clearTimeout(_ttTimer);
    tooltip.classList.remove('show');
  });
}

/* ================================================
   CELL NOTES
================================================ */
var _noteDialog = null;
var _noteCallback = null;

function initNoteDialog() {
  var dlg = document.createElement('div');
  dlg.id = 'noteDialog';
  dlg.style.cssText =
    'position:fixed;top:50%;left:50%;transform:translate(-50%,-50%);' +
    'z-index:999999;background:var(--bg-dropdown);border:1px solid var(--border);' +
    'border-radius:var(--radius-lg);box-shadow:var(--shadow-lg);' +
    'padding:16px;width:300px;display:none;flex-direction:column;gap:10px;';
  dlg.innerHTML =
    '<div style="font-size:12px;font-weight:600;color:var(--text-primary)">Nota para la celda</div>' +
    '<textarea id="noteTextarea" style="background:var(--bg-panel);border:1px solid var(--border);' +
    'border-radius:var(--radius);color:var(--text-primary);font-family:var(--font-ui);font-size:12px;' +
    'padding:6px 8px;outline:none;resize:vertical;min-height:80px;user-select:text;" placeholder="Escribe una nota..."></textarea>' +
    '<div style="display:flex;gap:6px;justify-content:flex-end;">' +
    '<button id="noteCancelBtn" class="btn btn-ghost btn-sm">Cancelar</button>' +
    '<button id="noteSaveBtn" class="btn btn-primary btn-sm">Guardar</button>' +
    '</div>';
  document.body.appendChild(dlg);
  _noteDialog = dlg;

  document.getElementById('noteCancelBtn').addEventListener('click', function() {
    dlg.style.display = 'none';
    _noteCallback = null;
  });
  document.getElementById('noteSaveBtn').addEventListener('click', function() {
    var val = document.getElementById('noteTextarea').value;
    dlg.style.display = 'none';
    if (_noteCallback) { _noteCallback(val); _noteCallback = null; }
  });
  document.getElementById('noteTextarea').addEventListener('keydown', function(e) {
    if (e.key === 'Enter' && e.ctrlKey) {
      var val2 = document.getElementById('noteTextarea').value;
      dlg.style.display = 'none';
      if (_noteCallback) { _noteCallback(val2); _noteCallback = null; }
    }
    if (e.key === 'Escape') { dlg.style.display = 'none'; _noteCallback = null; }
  });
}

function addNote(r, c) {
  if (!_noteDialog) return;
  var cur = (STATE.cellNotes ? STATE.cellNotes[r + ',' + c] : '') || '';
  document.getElementById('noteTextarea').value = cur;
  _noteDialog.style.display = 'flex';
  document.getElementById('noteTextarea').focus();
  _noteCallback = function(note) {
    if (!STATE.cellNotes) STATE.cellNotes = {};
    if (note.trim()) STATE.cellNotes[r + ',' + c] = note;
    else delete STATE.cellNotes[r + ',' + c];
    renderNoteIndicator(r, c);
    markDirty();
    showToast(note.trim() ? 'Nota guardada' : 'Nota eliminada', 'success');
  };
}

function renderNoteIndicator(r, c, td) {
  if (!td) td = getCellEl(r, c);
  if (!td) return;
  td.querySelectorAll('.note-indicator').forEach(function(el) { el.remove(); });
  var note = STATE.cellNotes ? STATE.cellNotes[r + ',' + c] : null;
  if (note) {
    var dot = document.createElement('span');
    dot.className = 'note-indicator';
    dot.style.cssText = 'position:absolute;top:2px;right:2px;width:5px;height:5px;' +
      'border-radius:50%;background:var(--formula);pointer-events:none;z-index:2;';
    td.appendChild(dot);
  }
}

/* ================================================
   RENAME FILE
================================================ */
function startRename() {
  var inp = document.getElementById('fileNameInput');
  var lbl = document.getElementById('fileName');
  inp.value = STATE.fileName;
  lbl.style.display = 'none';
  inp.style.display = '';
  document.getElementById('btnRename').style.display = 'none';
  inp.focus(); inp.select();
}
function commitRename() {
  var inp = document.getElementById('fileNameInput');
  var val = inp.value.trim();
  if (val) STATE.fileName = val;
  document.getElementById('fileName').textContent = STATE.fileName;
  document.getElementById('fileName').style.display = '';
  inp.style.display = 'none';
  document.getElementById('btnRename').style.display = '';
  markDirty();
}
function cancelRename() {
  document.getElementById('fileNameInput').style.display = 'none';
  document.getElementById('fileName').style.display = '';
  document.getElementById('btnRename').style.display = '';
}

/* ================================================
   FUNCIONES GLOBALES EXPORTADAS
================================================ */
window.loadFile = loadFile;
window.exportXLSX = exportXLSX;
window.exportCSV = exportCSV;
window.showToast = showToast;
window.showLoading = showLoading;
window.hideLoading = hideLoading;
window.markDirty = markDirty;
window.columnLabel = columnLabel;
window.getCellEl = getCellEl;
window.getCellSpan = getCellSpan;
window.getCellFmt = getCellFmt;
window.setCellFmt = setCellFmt;
window.getCellBorders = getCellBorders;
window.applyFmtToSpan = applyFmtToSpan;
window.applyBordersToTd = applyBordersToTd;
window.refreshCellFmt = refreshCellFmt;
window.applyCondFmtToCell = applyCondFmtToCell;
window.FMT = FMT;
window.BORDERS = BORDERS;
window.COND_RULES = COND_RULES;