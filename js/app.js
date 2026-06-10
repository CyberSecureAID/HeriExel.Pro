/* ===============================================
   HeriExcel.Pro — app.js  v7.12 (ESTILOS CORREGIDOS)
   Carga de archivos con estilos completos
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
   RENDER TABLE (sin cambios mayores)
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

/* ================================================
   FUNCIONES AUXILIARES (copiadas de app-core.js para completar)
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
   OTRAS FUNCIONES (mínimas para evitar errores)
================================================ */
function showToast(msg, type) {
  if (type === undefined) type = '';
  var el = document.getElementById('toast');
  if (!el) return;
  el.textContent = msg;
  el.className = 'toast ' + type + ' show';
  clearTimeout(window._toastTimer);
  window._toastTimer = setTimeout(function() { el.classList.remove('show'); }, 3000);
}

function showLoading(msg) {
  var ov = document.getElementById('loadingOverlay');
  if (!ov) {
    ov = document.createElement('div');
    ov.id = 'loadingOverlay';
    ov.className = 'loading-overlay';
    ov.innerHTML = '<div class="loading-spinner"></div><div class="loading-text" id="loadingText"></div>';
    document.body.appendChild(ov);
  }
  var lt = ov.querySelector('#loadingText');
  if (lt) lt.textContent = msg || 'Cargando...';
  ov.classList.remove('hidden');
}

function hideLoading() {
  var ov = document.getElementById('loadingOverlay');
  if (ov) ov.classList.add('hidden');
}

function getCellEl(r, c) {
  var tbody = document.getElementById('tableBody');
  if (!tbody) return null;
  var tr = tbody.querySelector('tr[data-row="' + r + '"]');
  if (!tr) return null;
  return tr.querySelector('td[data-col="' + c + '"]') || null;
}

function getCellSpan(r, c) {
  var td = getCellEl(r, c);
  return td ? td.querySelector('.cell-display') : null;
}

function columnLabel(index) {
  var label = '', i = index;
  do {
    label = String.fromCharCode(65 + (i % 26)) + label;
    i = Math.floor(i / 26) - 1;
  } while (i >= 0);
  return label;
}

function evaluateFormula(formula, r, c) {
  return formula.startsWith('=') ? '#ERROR!' : formula;
}

function formatCellValue(val, numFormat) {
  if (val === null || val === undefined) return '';
  if (!numFormat || numFormat === 'general' || numFormat === 'text') return val;
  var num = parseFloat(String(val).replace(/[^0-9.\-]/g, ''));
  if (isNaN(num)) return val;
  switch (numFormat) {
    case 'number': return num.toLocaleString('es-ES', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    case 'integer': return Math.round(num).toLocaleString('es-ES');
    case 'currency': return num.toLocaleString('es-ES', { style: 'currency', currency: 'USD', minimumFractionDigits: 2 });
    case 'percent': return num.toLocaleString('es-ES', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' %';
    case 'scientific': return num.toExponential(2).toUpperCase();
    case 'date': return val;
    default: return val;
  }
}

function fmtKey(r, c) { return STATE.activeSheet + ':' + r + ',' + c; }
function borderKey(r, c) { return STATE.activeSheet + ':' + r + ',' + c; }

function getCellFmt(r, c) {
  var k = fmtKey(r, c);
  if (!FMT[k]) FMT[k] = {};
  return FMT[k];
}

function setCellFmt(r, c, props) {
  var k = fmtKey(r, c);
  if (!FMT[k]) FMT[k] = {};
  Object.assign(FMT[k], props);
}

function getCellBorders(r, c) {
  var k = borderKey(r, c);
  if (!BORDERS[k]) BORDERS[k] = { top: null, bottom: null, left: null, right: null };
  if (!Object.prototype.hasOwnProperty.call(BORDERS[k], 'top')) BORDERS[k].top = null;
  if (!Object.prototype.hasOwnProperty.call(BORDERS[k], 'bottom')) BORDERS[k].bottom = null;
  if (!Object.prototype.hasOwnProperty.call(BORDERS[k], 'left')) BORDERS[k].left = null;
  if (!Object.prototype.hasOwnProperty.call(BORDERS[k], 'right')) BORDERS[k].right = null;
  return BORDERS[k];
}

function applyFmtToSpan(span, fmt) {
  if (!span || !fmt) return;
  span.style.fontWeight = fmt.bold ? '700' : '';
  span.style.fontStyle = fmt.italic ? 'italic' : '';
  var decorations = [];
  if (fmt.underline) decorations.push('underline');
  if (fmt.strike) decorations.push('line-through');
  span.style.textDecoration = decorations.join(' ') || '';
  span.style.fontFamily = fmt.font || '';
  span.style.fontSize = fmt.size || '';
  span.style.color = fmt.color || '';
  span.style.backgroundColor = fmt.bgColor || '';
  span.style.textAlign = fmt.align || '';

  var td = span.closest('td');
  if (td) {
    var ALL_STYLES = ['cell-style-good', 'cell-style-bad', 'cell-style-neutral', 'cell-style-header', 'cell-style-total', 'cell-style-warning'];
    ALL_STYLES.forEach(function(s) { td.classList.remove(s); });
    if (fmt.cellStyle && fmt.cellStyle !== 'normal') {
      td.classList.add('cell-style-' + fmt.cellStyle);
    }
  }
}

function applyBordersToTd(td, r, c) {
  if (!td) return;
  var b = getCellBorders(r, c);
  ['top', 'bottom', 'left', 'right'].forEach(function(side) {
    var bdef = b[side];
    var propName = 'border' + side.charAt(0).toUpperCase() + side.slice(1);
    td.style[propName] = bdef ? ((bdef.width || '1px') + ' ' + (bdef.style || 'solid') + ' ' + (bdef.color || '#7c8cf8')) : '';
  });
}

function refreshCellFmt(r, c) {
  var span = getCellSpan(r, c);
  if (span) applyFmtToSpan(span, getCellFmt(r, c));
  var td = getCellEl(r, c);
  if (td) applyBordersToTd(td, r, c);
}

function applyCondFmtToCell(span, r, c, val) {
  if (!span) return;
  var fmt = getCellFmt(r, c);
  span.style.backgroundColor = fmt.bgColor || '';
  span.style.color = fmt.color || '';
  span.style.fontWeight = fmt.bold ? '700' : '';
}

function applyAllCondRules() { /* placeholder */ }

function renderCondRuleList() { /* placeholder */ }

function selectEntireRow(r) { /* placeholder minimal */ }
function selectEntireCol(c) { /* placeholder minimal */ }

/* ================================================
   EXPORTAR FUNCIONES GLOBALES
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