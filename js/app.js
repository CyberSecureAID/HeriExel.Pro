/* ===============================================
   HeriExcel.Pro — app.js  v7.8
   Carga de archivos, renderizado de tabla,
   contexto, filtros, busqueda, zoom,
   exportacion, impresion, fijar filas,
   notas, fill-handle

   CAMBIOS v7.8:
   - exportXLSX() ahora exporta COLORES, FUENTES,
     BORDES, ALINEACIÓN y FORMATO NUMÉRICO usando
     la API de estilos de SheetJS (cell_styles).
   - Menú contextual: añadida opción
     "Eliminar columna" junto a "Eliminar fila".
   - Todos los fixes de v7.7 se mantienen intactos.
=============================================== */
'use strict';

/* ================================================
   INIT — único DOMContentLoaded
================================================ */
document.addEventListener('DOMContentLoaded', function() {
  initAutocomplete();
  initFileHandlers();
  initContextMenu();
  initFindReplace();
  initFilterDropdown();
  initRibbonActions();
  initMouseGlobalEvents();
  initTooltip();
  initFillHandle();
  initNoteDialog();
  renderCondRuleList();

  /* Módulos opcionales (date picker + eyedropper) */
  if (typeof initDatePicker === 'function') initDatePicker();
  if (typeof initEyeDropper === 'function') initEyeDropper();
});

/* ================================================
   FILE HANDLERS
================================================ */
function initFileHandlers() {
  var fileInput = $('fileInput');
  var dropZone  = $('dropZone');

  fileInput.addEventListener('change', function(e) {
    var file = e.target.files[0];
    if (file) loadFile(file);
    fileInput.value = '';
  });

  dropZone.addEventListener('dragover', function(e) {
    e.preventDefault();
    dropZone.classList.add('drag-over');
  });
  dropZone.addEventListener('dragleave', function() { dropZone.classList.remove('drag-over'); });
  dropZone.addEventListener('drop', function(e) {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    var file = e.dataTransfer.files[0];
    if (file) loadFile(file);
  });

  $('btnReset').addEventListener('click', resetApp);
  $('btnRename').addEventListener('click', startRename);
  $('fileNameInput').addEventListener('blur',    commitRename);
  $('fileNameInput').addEventListener('keydown', function(e) {
    if (e.key === 'Enter') commitRename();
    if (e.key === 'Escape') cancelRename();
  });
}

function loadFile(file) {
  STATE.fileName = file.name;
  showLoading('Leyendo archivo...');

  var reader = new FileReader();
  reader.onload = function(e) {
    try {
      var data = new Uint8Array(e.target.result);
      var wb   = XLSX.read(data, { type: 'array', cellDates: true });
      STATE.workbook   = wb;
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
  var ws  = STATE.workbook.Sheets[sheetName];
  STATE.wsRaw = ws;

  var jsonData = XLSX.utils.sheet_to_json(ws, {
    header: 1, raw: false, defval: '', blankrows: true,
  });

  /* Eliminar filas vacías al final */
  while (jsonData.length && jsonData[jsonData.length - 1].every(function(v) { return v === ''; })) {
    jsonData.pop();
  }

  STATE.data = jsonData.map(function(row) { return row.slice(); });

  if (!STATE.data.length) {
    STATE.data = [new Array(10).fill('')];
  }

  var maxCols = Math.max.apply(null, STATE.data.map(function(r) { return r.length; }).concat([1]));
  STATE.data = STATE.data.map(function(row) {
    while (row.length < maxCols) row.push('');
    return row;
  });

  /* Mapa de celdas combinadas */
  STATE.mergeMap = {};
  if (ws['!merges']) {
    ws['!merges'].forEach(function(m) {
      for (var r = m.s.r; r <= m.e.r; r++) {
        for (var c = m.s.c; c <= m.e.c; c++) {
          if (r === m.s.r && c === m.s.c) {
            STATE.mergeMap[r + ',' + c] = { rowspan: m.e.r - m.s.r + 1, colspan: m.e.c - m.s.c + 1 };
          } else {
            STATE.mergeMap[r + ',' + c] = 'skip';
          }
        }
      }
    });
  }

  STATE.activeCellR   = -1;
  STATE.activeCellC   = -1;
  STATE.selectedCells.clear();
  STATE.selectedRow   = -1;
  STATE.frozen        = false;
  STATE.activeFilters = {};
  STATE.undoStack     = [];
  STATE.redoStack     = [];
  STATE.dirty         = false;
  STATE.colWidths     = {};
  STATE.rowHeights    = {};

  renderSheetTabs();
  renderTable();
  updateStats();
  hideLoading();
  updateUnsaved();
}

/* ================================================
   SHOW / HIDE EDITOR
================================================ */
function showEditor() {
  $('uploadSection').classList.add('hidden');
  $('editorSection').classList.remove('hidden');
  $('headerCenter').style.display = '';
  $('btnReset').classList.remove('hidden');
  $('btnExportMenu').classList.remove('hidden');
  $('formatBar').classList.remove('hidden');
  $('fxBar').classList.remove('hidden');
  $('btnRename').style.display = '';
  $('fileName').textContent = STATE.fileName;
  updateStats();
}

function resetApp() {
  STATE.workbook = null; STATE.data = []; STATE.sheetNames = [];
  STATE.activeSheet = ''; STATE.dirty = false;
  STATE.activeCellR = -1; STATE.activeCellC = -1;
  STATE.selectedCells.clear();
  STATE.undoStack = []; STATE.redoStack = [];
  STATE.colWidths = {}; STATE.rowHeights = {};
  $('uploadSection').classList.remove('hidden');
  $('editorSection').classList.add('hidden');
  $('headerCenter').style.display = 'none';
  $('btnReset').classList.add('hidden');
  $('btnExportMenu').classList.add('hidden');
  $('formatBar').classList.add('hidden');
  $('fxBar').classList.add('hidden');
  $('tableHead').innerHTML = '';
  $('tableBody').innerHTML = '';
  $('sheetTabs').innerHTML = '';
  $('fileName').textContent = '—';
  if (typeof dpHide === 'function') dpHide();
}

/* ================================================
   SHEET TABS
   B-38 FIX: try/catch
================================================ */
function renderSheetTabs() {
  var container = $('sheetTabs');
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
   RENDER TABLE
================================================ */
function renderTable() {
  var thead = $('tableHead');
  var tbody = $('tableBody');
  thead.innerHTML = '';
  tbody.innerHTML = '';

  if (!STATE.data.length) return;

  var numCols = STATE.data[0] ? STATE.data[0].length : 0;

  /* HEAD */
  var headRow  = document.createElement('tr');
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
      label.title       = headerVal;

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
        th.style.width    = STATE.colWidths[c] + 'px';
      }

      headRow.appendChild(th);
    })(ci);
  }
  thead.appendChild(headRow);

  /* BODY */
  var startRow = STATE.frozen ? 1 : 0;
  var frag     = document.createDocumentFragment();

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
          var fmt    = getCellFmt(r, c);

          var span = document.createElement('span');
          span.className = 'cell-display';

          var displayVal = rawVal;
          var isFormula  = typeof rawVal === 'string' && rawVal.startsWith('=');
          if (isFormula) {
            displayVal = evaluateFormula(rawVal, r, c);
            var isErr  = displayVal === '#ERROR!' || displayVal === '#DIV/0!' || displayVal === '#N/A' || displayVal === '#NUM!';
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
            td.style.width    = STATE.colWidths[c] + 'px';
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
              STATE.fillStart      = { r: r, c: c };
              STATE.fillEnd        = { r: r, c: c };
              STATE.fillAnchorR    = r;
              STATE.fillAnchorC    = c;
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
          td.addEventListener('dblclick',  function() { enterEdit(r, c); });
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

function isRowHidden(r) {
  if (!Object.keys(STATE.activeFilters).length) return false;
  for (var colStr in STATE.activeFilters) {
    var c   = parseInt(colStr);
    var val = String((STATE.data[r] ? STATE.data[r][c] : '') || '');
    if (STATE.activeFilters[colStr].indexOf(val) === -1) return true;
  }
  return false;
}

/* ================================================
   FILL RANGE HIGHLIGHT
================================================ */
function highlightFillRange(start, end) {
  if (!start || !end) return;
  var tbody = $('tableBody'); if (!tbody) return;
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

/* ================================================
   FREEZE
================================================ */
function updateFreezeLine() {
  var line = document.querySelector('.freeze-col-line');
  if (!STATE.frozen) { if (line) line.remove(); return; }
}

/* ================================================
   STATS
================================================ */
function updateStats() {
  var rows = STATE.data.length;
  var cols = STATE.data[0] ? STATE.data[0].length : 0;
  $('statRows').textContent = rows + ' filas';
  $('statCols').textContent = cols + ' columnas';
}

/* ================================================
   DIRTY / UNSAVED
================================================ */
function markDirty() {
  STATE.dirty = true;
  updateUnsaved();
}
function updateUnsaved() {
  var dot  = $('fileDot');
  var sep  = $('unsavedSep');
  var stat = $('statUnsaved');
  var show = STATE.dirty;
  if (dot)  dot.classList.toggle('visible', show);
  if (sep)  sep.style.display  = show ? '' : 'none';
  if (stat) stat.style.display = show ? '' : 'none';
}

/* ================================================
   ZOOM
================================================ */
function applyZoom() {
  var table  = $('excelTable');
  var scroll = $('tableScroll');
  if (!table || !scroll) return;

  var scale = STATE.zoom / 100;

  table.style.transform = 'scale(1)';
  table.style.transformOrigin = 'top left';

  var naturalW = table.scrollWidth;
  var naturalH = table.scrollHeight;

  table.style.transform = 'scale(' + scale + ')';

  scroll.style.paddingRight  = Math.max(0, naturalW * scale - scroll.clientWidth  + 20) + 'px';
  scroll.style.paddingBottom = Math.max(0, naturalH * scale - scroll.clientHeight + 20) + 'px';

  $('zoomIndicator').textContent = STATE.zoom + '%';
}

/* ================================================
   SORT
================================================ */
function sortByCol(asc) {
  var c = STATE.activeCellC;
  if (c < 0) return showToast('Haz clic en una columna primero', '');
  pushUndo();

  var startRow  = STATE.frozen ? 1 : 0;
  var header    = STATE.frozen ? [STATE.data[0].slice()] : [];
  var body      = STATE.data.slice(startRow).map(function(r) { return r.slice(); });
  var numCols   = STATE.data[0] ? STATE.data[0].length : 0;
  var sheet     = STATE.activeSheet;

  var origOrder = body.map(function(_, i) { return startRow + i; });

  var indexed = body.map(function(row, i) { return { row: row, origIdx: origOrder[i] }; });
  indexed.sort(function(a, b) {
    var va = a.row[c] || '', vb = b.row[c] || '';
    var na = parseFloat(va), nb = parseFloat(vb);
    if (!isNaN(na) && !isNaN(nb)) return asc ? na - nb : nb - na;
    return asc ? String(va).localeCompare(String(vb)) : String(vb).localeCompare(String(va));
  });

  var newFmt     = {}, newBorders = {};

  indexed.forEach(function(item, newBodyIdx) {
    var newR  = startRow + newBodyIdx;
    var origR = item.origIdx;
    for (var col = 0; col < numCols; col++) {
      var oldKey = sheet + ':' + origR + ',' + col;
      var newKey = sheet + ':' + newR  + ',' + col;
      if (FMT[oldKey])     newFmt[newKey]     = Object.assign({}, FMT[oldKey]);
      if (BORDERS[oldKey]) newBorders[newKey] = Object.assign({}, BORDERS[oldKey]);
    }
  });

  if (STATE.frozen) {
    for (var col2 = 0; col2 < numCols; col2++) {
      var hKey = sheet + ':0,' + col2;
      if (FMT[hKey])     newFmt[hKey]     = Object.assign({}, FMT[hKey]);
      if (BORDERS[hKey]) newBorders[hKey] = Object.assign({}, BORDERS[hKey]);
    }
  }

  for (var col3 = 0; col3 < numCols; col3++) {
    for (var r3 = startRow; r3 < STATE.data.length; r3++) {
      var k3 = sheet + ':' + r3 + ',' + col3;
      delete FMT[k3];
      delete BORDERS[k3];
    }
  }
  Object.assign(FMT,     newFmt);
  Object.assign(BORDERS, newBorders);

  STATE.data = header.concat(indexed.map(function(item) { return item.row; }));
  renderTable(); markDirty();
  showToast(asc ? 'Ordenado A a Z' : 'Ordenado Z a A', 'success');
}

/* ================================================
   CLEAR CELLS
================================================ */
function clearActiveCells() {
  var cells = getTargetCells();
  if (!cells.length) return;
  pushUndo();
  cells.forEach(function(rc) {
    if (STATE.data[rc[0]]) STATE.data[rc[0]][rc[1]] = '';
    var span = getCellSpan(rc[0], rc[1]);
    if (span) { span.textContent = ''; span.classList.remove('formula-error'); }
  });
  $('fxInput').value = '';
  markDirty();
}

/* ================================================
   HELPERS: migración FMT/BORDERS
================================================ */
function shiftRowKeys(fromRow, delta, numCols) {
  var sheet = STATE.activeSheet;
  numCols   = numCols || (STATE.data[0] ? STATE.data[0].length : 0);

  if (delta > 0) {
    var maxRow = STATE.data.length + delta;
    for (var r = maxRow; r >= fromRow + delta; r--) {
      var origR = r - delta;
      for (var c = 0; c < numCols; c++) {
        var oldKey = sheet + ':' + origR + ',' + c;
        var newKey = sheet + ':' + r     + ',' + c;
        if (FMT[oldKey])     { FMT[newKey]     = Object.assign({}, FMT[oldKey]);     delete FMT[oldKey]; }     else { delete FMT[newKey]; }
        if (BORDERS[oldKey]) { BORDERS[newKey] = Object.assign({}, BORDERS[oldKey]); delete BORDERS[oldKey]; } else { delete BORDERS[newKey]; }
      }
    }
  } else {
    var deletedCount = -delta;
    var totalRows    = STATE.data.length;

    var savedFmt     = {}, savedBorders = {};
    for (var rs = fromRow + deletedCount; rs < totalRows; rs++) {
      for (var cs = 0; cs < numCols; cs++) {
        var srcKey2 = sheet + ':' + rs + ',' + cs;
        if (FMT[srcKey2])     savedFmt[rs + ',' + cs]     = Object.assign({}, FMT[srcKey2]);
        if (BORDERS[srcKey2]) savedBorders[rs + ',' + cs] = Object.assign({}, BORDERS[srcKey2]);
      }
    }

    for (var rd = fromRow; rd < totalRows; rd++) {
      for (var cd = 0; cd < numCols; cd++) {
        var dk = sheet + ':' + rd + ',' + cd;
        delete FMT[dk];
        delete BORDERS[dk];
      }
    }

    for (var rw = fromRow + deletedCount; rw < totalRows; rw++) {
      var newR2 = rw - deletedCount;
      for (var cw = 0; cw < numCols; cw++) {
        var savedFmtKey = rw + ',' + cw;
        var dstKey2     = sheet + ':' + newR2 + ',' + cw;
        if (savedFmt[savedFmtKey])     FMT[dstKey2]     = savedFmt[savedFmtKey];
        if (savedBorders[savedFmtKey]) BORDERS[dstKey2] = savedBorders[savedFmtKey];
      }
    }
  }
}

function shiftColKeys(fromCol, delta) {
  var sheet   = STATE.activeSheet;
  var numRows = STATE.data.length;

  if (delta > 0) {
    var maxCol = (STATE.data[0] ? STATE.data[0].length : 0) + delta;
    for (var c = maxCol; c >= fromCol + delta; c--) {
      var origC = c - delta;
      for (var r = 0; r < numRows; r++) {
        var oldKey = sheet + ':' + r + ',' + origC;
        var newKey = sheet + ':' + r + ',' + c;
        if (FMT[oldKey])     { FMT[newKey]     = Object.assign({}, FMT[oldKey]);     delete FMT[oldKey]; }     else { delete FMT[newKey]; }
        if (BORDERS[oldKey]) { BORDERS[newKey] = Object.assign({}, BORDERS[oldKey]); delete BORDERS[oldKey]; } else { delete BORDERS[newKey]; }
      }
    }
  } else {
    var deletedCount2 = -delta;
    var totalCols     = STATE.data[0] ? STATE.data[0].length : 0;

    var savedFmt2     = {}, savedBorders2 = {};
    for (var rs2 = 0; rs2 < numRows; rs2++) {
      for (var cs2 = fromCol + deletedCount2; cs2 < totalCols; cs2++) {
        var srcKey3 = sheet + ':' + rs2 + ',' + cs2;
        if (FMT[srcKey3])     savedFmt2[rs2 + ',' + cs2]     = Object.assign({}, FMT[srcKey3]);
        if (BORDERS[srcKey3]) savedBorders2[rs2 + ',' + cs2] = Object.assign({}, BORDERS[srcKey3]);
      }
    }

    for (var rd2 = 0; rd2 < numRows; rd2++) {
      for (var cd2 = fromCol; cd2 < totalCols; cd2++) {
        var dk2 = sheet + ':' + rd2 + ',' + cd2;
        delete FMT[dk2];
        delete BORDERS[dk2];
      }
    }

    for (var rw2 = 0; rw2 < numRows; rw2++) {
      for (var cw2 = fromCol + deletedCount2; cw2 < totalCols; cw2++) {
        var newC    = cw2 - deletedCount2;
        var savedK2 = rw2 + ',' + cw2;
        var dstKey3 = sheet + ':' + rw2 + ',' + newC;
        if (savedFmt2[savedK2])     FMT[dstKey3]     = savedFmt2[savedK2];
        if (savedBorders2[savedK2]) BORDERS[dstKey3] = savedBorders2[savedK2];
      }
    }
  }
}

/* ================================================
   HELPER: leer cantidad del selector en el dropdown
================================================ */
function getQtyFromInput(inputId, maxVal) {
  var el = document.getElementById(inputId);
  if (!el) return 1;
  var v = parseInt(el.value, 10);
  if (isNaN(v) || v < 1) return 1;
  if (maxVal && v > maxVal) return maxVal;
  return v;
}

/* ================================================
   ADD / DELETE ROWS & COLS
================================================ */
function addRowAtEnd() {
  closeAllDropdowns(null);
  var qty  = getQtyFromInput('qtyAddRowEnd', 500);
  pushUndo();
  var cols = STATE.data[0] ? STATE.data[0].length : 1;
  for (var i = 0; i < qty; i++) {
    STATE.data.push(new Array(cols).fill(''));
  }
  renderTable(); updateStats(); markDirty();
  showToast(qty === 1 ? 'Fila agregada al final' : qty + ' filas agregadas al final', 'success');
}

function addRowAbove() {
  closeAllDropdowns(null);
  var r = STATE.activeCellR;
  if (r < 0) return showToast('Selecciona una celda primero', '');
  var qty  = getQtyFromInput('qtyAddRowAbove', 500);
  pushUndo();
  var cols = STATE.data[0] ? STATE.data[0].length : 1;
  shiftRowKeys(r, +qty, cols);
  for (var i = 0; i < qty; i++) {
    STATE.data.splice(r, 0, new Array(cols).fill(''));
  }
  STATE.activeCellR = r + qty;
  renderTable(); updateStats(); markDirty();
  showToast(qty === 1 ? 'Fila insertada arriba' : qty + ' filas insertadas arriba', 'success');
}

function addRowBelow() {
  closeAllDropdowns(null);
  var r = STATE.activeCellR;
  if (r < 0) return showToast('Selecciona una celda primero', '');
  var qty  = getQtyFromInput('qtyAddRowBelow', 500);
  pushUndo();
  var cols = STATE.data[0] ? STATE.data[0].length : 1;
  shiftRowKeys(r + 1, +qty, cols);
  for (var i = 0; i < qty; i++) {
    STATE.data.splice(r + 1 + i, 0, new Array(cols).fill(''));
  }
  renderTable(); updateStats(); markDirty();
  showToast(qty === 1 ? 'Fila insertada abajo' : qty + ' filas insertadas abajo', 'success');
}

function duplicateRow() {
  closeAllDropdowns(null);
  var r = STATE.activeCellR;
  if (r < 0) return showToast('Selecciona una celda primero', '');
  pushUndo();
  var cols  = STATE.data[0] ? STATE.data[0].length : 1;
  var sheet = STATE.activeSheet;
  shiftRowKeys(r + 1, +1, cols);
  for (var c = 0; c < cols; c++) {
    var srcKey = sheet + ':' + r       + ',' + c;
    var dstKey = sheet + ':' + (r + 1) + ',' + c;
    if (FMT[srcKey])     FMT[dstKey]     = Object.assign({}, FMT[srcKey]);
    if (BORDERS[srcKey]) BORDERS[dstKey] = Object.assign({}, BORDERS[srcKey]);
  }
  STATE.data.splice(r + 1, 0, STATE.data[r].slice());
  renderTable(); updateStats(); markDirty();
  showToast('Fila duplicada', 'success');
}

function deleteActiveRow() {
  closeAllDropdowns(null);
  var r = STATE.activeCellR;
  if (r < 0) return showToast('Selecciona una celda primero', '');
  if (STATE.data.length <= 1) return showToast('No se puede eliminar la única fila', '');
  pushUndo();
  var cols = STATE.data[0] ? STATE.data[0].length : 1;
  shiftRowKeys(r, -1, cols);
  STATE.data.splice(r, 1);
  STATE.activeCellR = Math.min(r, STATE.data.length - 1);
  STATE.selectedCells.clear();
  renderTable(); updateStats(); markDirty();
  showToast('Fila eliminada', 'info');
}

function deleteSelectionRows() {
  closeAllDropdowns(null);
  if (!STATE.selectedCells.size) return showToast('Selecciona celdas primero', '');
  var rows = Array.from(new Set(Array.from(STATE.selectedCells).map(function(k) { return parseInt(k.split(',')[0]); })));
  rows.sort(function(a, b) { return b - a; });
  if (STATE.data.length <= rows.length) return showToast('No se pueden eliminar todas las filas', '');
  pushUndo();
  var cols = STATE.data[0] ? STATE.data[0].length : 1;
  rows.forEach(function(r) {
    shiftRowKeys(r, -1, cols);
    STATE.data.splice(r, 1);
  });
  STATE.activeCellR = -1; STATE.selectedCells.clear();
  renderTable(); updateStats(); markDirty();
  showToast(rows.length + ' fila(s) eliminadas', 'info');
}

function addColAtEnd() {
  closeAllDropdowns(null);
  var qty = getQtyFromInput('qtyAddColEnd', 100);
  pushUndo();
  for (var i = 0; i < qty; i++) {
    STATE.data.forEach(function(row) { row.push(''); });
  }
  renderTable(); updateStats(); markDirty();
  showToast(qty === 1 ? 'Columna agregada al final' : qty + ' columnas agregadas al final', 'success');
}

function addColBefore() {
  closeAllDropdowns(null);
  var c = STATE.activeCellC;
  if (c < 0) return showToast('Selecciona una celda primero', '');
  var qty = getQtyFromInput('qtyAddColBefore', 100);
  pushUndo();
  shiftColKeys(c, +qty);
  for (var i = 0; i < qty; i++) {
    STATE.data.forEach(function(row) { row.splice(c + i, 0, ''); });
  }
  var newWidths = {};
  Object.keys(STATE.colWidths).forEach(function(k) {
    var ki = parseInt(k);
    if (ki >= c) newWidths[ki + qty] = STATE.colWidths[k];
    else         newWidths[ki]       = STATE.colWidths[k];
  });
  STATE.colWidths = newWidths;
  renderTable(); updateStats(); markDirty();
  showToast(qty === 1 ? 'Columna insertada antes' : qty + ' columnas insertadas antes', 'success');
}

function splitColumnInHalf() {
  closeAllDropdowns(null);
  var c = STATE.activeCellC;
  if (c < 0) return showToast('Selecciona una columna primero', '');
  pushUndo();
  shiftColKeys(c + 1, +1);
  STATE.data.forEach(function(row) { row.splice(c + 1, 0, ''); });
  var newWidths = {};
  Object.keys(STATE.colWidths).forEach(function(k) {
    var ki = parseInt(k);
    if (ki > c) newWidths[ki + 1] = STATE.colWidths[k];
    else        newWidths[ki]     = STATE.colWidths[k];
  });
  var currentWidth = STATE.colWidths[c] || 120;
  var half         = Math.max(60, Math.floor(currentWidth / 2));
  newWidths[c]     = half;
  newWidths[c + 1] = half;
  STATE.colWidths  = newWidths;
  renderTable(); updateStats(); markDirty();
  showToast('Columna ' + columnLabel(c) + ' dividida en ' + columnLabel(c) + ' y ' + columnLabel(c + 1), 'success');
}

function deleteActiveCol() {
  closeAllDropdowns(null);
  var c = STATE.activeCellC;
  if (c < 0) return showToast('Selecciona una celda primero', '');
  if ((STATE.data[0] ? STATE.data[0].length : 0) <= 1) return showToast('No se puede eliminar la única columna', '');
  pushUndo();
  shiftColKeys(c, -1);
  STATE.data.forEach(function(row) { row.splice(c, 1); });
  var newWidths2 = {};
  Object.keys(STATE.colWidths).forEach(function(k) {
    var ki = parseInt(k);
    if (ki < c)      newWidths2[ki]     = STATE.colWidths[k];
    else if (ki > c) newWidths2[ki - 1] = STATE.colWidths[k];
  });
  STATE.colWidths   = newWidths2;
  STATE.activeCellC = Math.max(0, c - 1);
  renderTable(); updateStats(); markDirty();
  showToast('Columna eliminada', 'info');
}

/* ================================================
   ESTILOS PARA LOS SELECTORES DE CANTIDAD
================================================ */
function injectQtyStyles() {
  if (document.getElementById('hx-qty-styles')) return;
  var style = document.createElement('style');
  style.id = 'hx-qty-styles';
  style.textContent = [
    '.hx-qty-row {',
    '  display: flex;',
    '  align-items: center;',
    '  gap: 6px;',
    '  padding: 2px 10px 6px;',
    '  border-top: 1px solid var(--border);',
    '  margin-top: 2px;',
    '}',
    '.hx-qty-label {',
    '  font-size: 10.5px;',
    '  color: var(--text-muted);',
    '  white-space: nowrap;',
    '  flex-shrink: 0;',
    '}',
    '.hx-qty-input {',
    '  width: 60px;',
    '  height: 24px;',
    '  background: var(--bg-panel);',
    '  border: 1px solid var(--border);',
    '  border-radius: var(--radius);',
    '  color: var(--text-primary);',
    '  font-family: var(--font-mono);',
    '  font-size: 12px;',
    '  padding: 0 6px;',
    '  outline: none;',
    '  text-align: center;',
    '  user-select: text;',
    '}',
    '.hx-qty-input:focus { border-color: var(--accent); }',
    '.hx-qty-hint {',
    '  font-size: 10px;',
    '  color: var(--text-muted);',
    '  opacity: 0.7;',
    '}',
  ].join('\n');
  document.head.appendChild(style);
}

/* ================================================
   RIBBON ACTIONS
================================================ */
function initRibbonActions() {
  injectQtyStyles();

  _injectQtyRow('ddRowsMenu', 'btnAddRow',     'qtyAddRowEnd',   500, 'Cantidad de filas a agregar al final');
  _injectQtyRow('ddRowsMenu', 'btnAddRowAbove','qtyAddRowAbove', 500, 'Cantidad de filas a insertar arriba');
  _injectQtyRow('ddRowsMenu', 'btnAddRowBelow','qtyAddRowBelow', 500, 'Cantidad de filas a insertar abajo');

  _injectQtyRow('ddColsMenu', 'btnAddCol',      'qtyAddColEnd',    100, 'Cantidad de columnas a agregar al final');
  _injectQtyRow('ddColsMenu', 'btnAddColBefore','qtyAddColBefore', 100, 'Cantidad de columnas a insertar antes');

  /* Filas */
  var btnAddRow              = $('btnAddRow');
  var btnAddRowAbove         = $('btnAddRowAbove');
  var btnAddRowBelow         = $('btnAddRowBelow');
  var btnDuplicateRow        = $('btnDuplicateRow');
  var btnDeleteRow           = $('btnDeleteRow');
  var btnDeleteSelectionRows = $('btnDeleteSelectionRows');
  if (btnAddRow)              btnAddRow.addEventListener('click',              addRowAtEnd);
  if (btnAddRowAbove)         btnAddRowAbove.addEventListener('click',         addRowAbove);
  if (btnAddRowBelow)         btnAddRowBelow.addEventListener('click',         addRowBelow);
  if (btnDuplicateRow)        btnDuplicateRow.addEventListener('click',        duplicateRow);
  if (btnDeleteRow)           btnDeleteRow.addEventListener('click',           deleteActiveRow);
  if (btnDeleteSelectionRows) btnDeleteSelectionRows.addEventListener('click', deleteSelectionRows);

  /* Columnas */
  var btnAddCol       = $('btnAddCol');
  var btnAddColBefore = $('btnAddColBefore');
  var btnSplitCol     = $('btnSplitCol');
  var btnDeleteCol    = $('btnDeleteCol');
  if (btnAddCol)       btnAddCol.addEventListener('click',       addColAtEnd);
  if (btnAddColBefore) btnAddColBefore.addEventListener('click', addColBefore);
  if (btnSplitCol)     btnSplitCol.addEventListener('click',    splitColumnInHalf);
  if (btnDeleteCol)    btnDeleteCol.addEventListener('click',   deleteActiveCol);

  /* Exportar */
  var btnExportXLSX = $('btnExportXLSX');
  var btnExportCSV  = $('btnExportCSV');
  if (btnExportXLSX) btnExportXLSX.addEventListener('click', function() { closeAllDropdowns(null); exportXLSX(); });
  if (btnExportCSV)  btnExportCSV.addEventListener('click',  function() { closeAllDropdowns(null); exportCSV();  });

  /* Ordenar */
  var btnSortAsc  = $('btnSortAsc');
  var btnSortDesc = $('btnSortDesc');
  if (btnSortAsc)  btnSortAsc.addEventListener('click',  function() { sortByCol(true);  });
  if (btnSortDesc) btnSortDesc.addEventListener('click', function() { sortByCol(false); });

  /* Congelar */
  var btnFreeze = $('btnFreeze');
  if (btnFreeze) {
    btnFreeze.addEventListener('click', function() {
      if (!STATE.data.length) return showToast('No hay datos cargados', '');
      STATE.frozen = !STATE.frozen;
      btnFreeze.classList.toggle('active', STATE.frozen);
      renderTable();
      showToast(STATE.frozen ? 'Encabezado fijado' : 'Encabezado liberado', 'info');
    });
  }

  /* Limpiar celda */
  var btnClearCell = $('btnClearCell');
  if (btnClearCell) btnClearCell.addEventListener('click', clearActiveCells);

  /* Buscar inline */
  var searchInput = $('searchInput');
  if (searchInput) {
    searchInput.addEventListener('input', function() {
      var q = searchInput.value.trim();
      clearSearchHighlights();
      if (!q) { $('searchCountBadge').textContent = ''; return; }
      performSearch(q);
    });
  }

  /* Reemplazar inline */
  var btnReplace    = $('btnReplace');
  var btnReplaceAll = $('btnReplaceAll');

  if (btnReplace) {
    btnReplace.addEventListener('click', function() {
      var q   = $('searchInput').value.trim();
      var rep = $('replaceInput').value;
      if (!q || !_searchMatches.length) return showToast('Busca algo primero', '');
      var idx2 = _searchIdx < 0 ? 0 : _searchIdx;
      var cur  = _searchMatches[idx2]; if (!cur) return;
      pushUndo();
      var r    = cur.r, c = cur.c;
      if (r < 0 || c < 0) return;
      var esc   = q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      var regex = new RegExp(esc, 'gi');
      STATE.data[r][c] = String(STATE.data[r][c]).replace(regex, rep);
      var span = getCellSpan(r, c);
      if (span) {
        span.textContent = formatCellValue(STATE.data[r][c], getCellFmt(r, c).numFormat);
        span.classList.remove('search-match', 'search-current');
      }
      _searchMatches.splice(idx2, 1);
      $('searchCountBadge').textContent = _searchMatches.length || '';
      markDirty();
      showToast('Reemplazado', 'success');
    });
  }

  if (btnReplaceAll) {
    btnReplaceAll.addEventListener('click', function() {
      var q   = $('searchInput').value.trim();
      var rep = $('replaceInput').value;
      if (!q) return showToast('Escribe un término a buscar', '');
      pushUndo();
      var count = 0;
      var esc   = q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      var regex = new RegExp(esc, 'gi');
      STATE.data.forEach(function(row, r) {
        row.forEach(function(val, c) {
          if (String(val).toLowerCase().indexOf(q.toLowerCase()) !== -1) {
            STATE.data[r][c] = String(val).replace(regex, rep);
            count++;
          }
        });
      });
      clearSearchHighlights();
      renderTable();
      markDirty();
      showToast(count + ' reemplazo(s) realizados', 'success');
      closeAllDropdowns(null);
    });
  }

  /* Zoom */
  var btnZoomIn  = $('btnZoomIn');
  var btnZoomOut = $('btnZoomOut');
  if (btnZoomIn) {
    btnZoomIn.addEventListener('click', function() {
      if (STATE.zoom >= 200) return;
      STATE.zoom = Math.min(200, STATE.zoom + 10);
      applyZoom();
    });
  }
  if (btnZoomOut) {
    btnZoomOut.addEventListener('click', function() {
      if (STATE.zoom <= 50) return;
      STATE.zoom = Math.max(50, STATE.zoom - 10);
      applyZoom();
    });
  }
}

/* ================================================
   _injectQtyRow
================================================ */
function _injectQtyRow(menuId, afterId, inputId, maxVal, hint) {
  var menu = document.getElementById(menuId);
  var btn  = document.getElementById(afterId);
  if (!menu || !btn) return;

  var row = document.createElement('div');
  row.className = 'hx-qty-row';
  row.innerHTML =
    '<span class="hx-qty-label">Cantidad:</span>' +
    '<input id="' + inputId + '" class="hx-qty-input" type="number" ' +
    '  min="1" max="' + maxVal + '" value="1" ' +
    '  title="' + hint + '" autocomplete="off" />' +
    '<span class="hx-qty-hint">máx ' + maxVal + '</span>';

  row.addEventListener('click',      function(e) { e.stopPropagation(); });
  row.addEventListener('mousedown',  function(e) { e.stopPropagation(); });
  row.addEventListener('keydown',    function(e) { e.stopPropagation(); });

  btn.insertAdjacentElement('afterend', row);
}

/* ================================================
   EXPORT — v7.8: con estilos completos
   Convierte FMT y BORDERS a estilos XLSX nativos
   (font, fill, alignment, border, numFmt).
================================================ */

/**
 * Convierte un color CSS (#rrggbb / rgb(...) / nombre)
 * al formato ARGB que usa SheetJS: "FF" + rrggbb en mayúsculas.
 */
function _cssColorToArgb(css) {
  if (!css) return null;
  css = css.trim();
  /* #rgb → #rrggbb */
  if (/^#[0-9a-f]{3}$/i.test(css)) {
    css = '#' + css[1]+css[1] + css[2]+css[2] + css[3]+css[3];
  }
  /* #rrggbb */
  if (/^#[0-9a-f]{6}$/i.test(css)) {
    return 'FF' + css.slice(1).toUpperCase();
  }
  /* rgb(r,g,b) */
  var m = css.match(/^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i);
  if (m) {
    return 'FF' +
      parseInt(m[1]).toString(16).padStart(2,'0').toUpperCase() +
      parseInt(m[2]).toString(16).padStart(2,'0').toUpperCase() +
      parseInt(m[3]).toString(16).padStart(2,'0').toUpperCase();
  }
  return null;
}

/**
 * Construye el objeto de estilo SheetJS para una celda dada (r, c).
 * Retorna null si no hay ningún estilo aplicado.
 */
function _buildCellStyle(r, c) {
  var fmtKey  = STATE.activeSheet + ':' + r + ',' + c;
  var fmt     = FMT[fmtKey]     || {};
  var borders = BORDERS[fmtKey] || {};

  var hasStyle =
    fmt.bold || fmt.italic || fmt.underline || fmt.strike ||
    fmt.font || fmt.size || fmt.color || fmt.bgColor ||
    fmt.align || fmt.numFormat ||
    fmt.cellStyle ||
    (borders.top && borders.top.color) ||
    (borders.bottom && borders.bottom.color) ||
    (borders.left && borders.left.color) ||
    (borders.right && borders.right.color);

  if (!hasStyle) return null;

  var style = {};

  /* ── FONT ── */
  var fontObj = {};
  if (fmt.bold)      fontObj.bold      = true;
  if (fmt.italic)    fontObj.italic    = true;
  if (fmt.underline) fontObj.underline = true;
  if (fmt.strike)    fontObj.strike    = true;
  if (fmt.font)      fontObj.name      = fmt.font.replace(/'/g,'').split(',')[0].trim();
  if (fmt.size) {
    var pts = parseFloat(fmt.size);
    if (!isNaN(pts)) fontObj.sz = pts * 0.75; /* px → pt aprox */
  }
  if (fmt.color) {
    var fc = _cssColorToArgb(fmt.color);
    if (fc) fontObj.color = { rgb: fc };
  }
  if (Object.keys(fontObj).length) style.font = fontObj;

  /* ── FILL (background color) ── */
  var bgColor = fmt.bgColor;
  /* Estilos predefinidos: derivar color de fondo */
  if (!bgColor && fmt.cellStyle) {
    var styleMap = {
      header:  '#1e1f2e',
      total:   '#2d3050',
      good:    'rgba(74,222,128,0.18)',
      bad:     'rgba(248,113,113,0.18)',
      neutral: 'rgba(251,191,36,0.18)',
      warning: 'rgba(251,146,60,0.18)',
    };
    bgColor = styleMap[fmt.cellStyle] || null;
  }
  if (bgColor) {
    var bc = _cssColorToArgb(bgColor);
    if (bc) {
      style.fill = { patternType: 'solid', fgColor: { rgb: bc } };
    }
  }

  /* ── ALIGNMENT ── */
  var alignObj = {};
  if (fmt.align) {
    var xlAlign = { left: 'left', center: 'center', right: 'right' };
    if (xlAlign[fmt.align]) alignObj.horizontal = xlAlign[fmt.align];
  }
  if (Object.keys(alignObj).length) style.alignment = alignObj;

  /* ── NUMBER FORMAT ── */
  var numFmtMap = {
    number:     '#,##0.00',
    integer:    '#,##0',
    currency:   '"$"#,##0.00',
    percent:    '0.00"%"',
    scientific: '0.00E+00',
    date:       'dd/mm/yyyy',
  };
  if (fmt.numFormat && numFmtMap[fmt.numFormat]) {
    style.numFmt = numFmtMap[fmt.numFormat];
  }

  /* ── BORDERS ── */
  var borderObj = {};
  var SIDES = ['top', 'bottom', 'left', 'right'];
  SIDES.forEach(function(side) {
    var bd = borders[side];
    if (!bd) return;
    var bStyle = bd.style || 'thin';
    /* Mapear estilos CSS a SheetJS */
    var xlStyle = 'thin';
    if (bStyle === 'dashed')  xlStyle = 'dashed';
    if (bStyle === 'double')  xlStyle = 'double';
    if (bd.width && parseFloat(bd.width) >= 2) xlStyle = 'medium';
    if (bd.width && parseFloat(bd.width) >= 3) xlStyle = 'thick';
    var bColor = _cssColorToArgb(bd.color);
    var entry  = { style: xlStyle };
    if (bColor) entry.color = { rgb: bColor };
    borderObj[side] = entry;
  });
  if (Object.keys(borderObj).length) style.border = borderObj;

  return Object.keys(style).length ? style : null;
}

/**
 * Exporta el libro con todos los estilos (colores, fuentes, bordes).
 * Usa la API de estilos de SheetJS (requiere la versión full incluida).
 */
function exportXLSX() {
  if (!STATE.data.length) return showToast('No hay datos para exportar', '');
  showLoading('Exportando...');
  setTimeout(function() {
    try {
      var wb = XLSX.utils.book_new();
      var ws = XLSX.utils.aoa_to_sheet(STATE.data);

      /* ── Aplicar estilos celda a celda ── */
      var numRows = STATE.data.length;
      var numCols = STATE.data[0] ? STATE.data[0].length : 0;

      for (var r = 0; r < numRows; r++) {
        for (var c = 0; c < numCols; c++) {
          var cellAddr = XLSX.utils.encode_cell({ r: r, c: c });
          var cellStyle = _buildCellStyle(r, c);
          if (cellStyle) {
            if (!ws[cellAddr]) {
              ws[cellAddr] = { t: 's', v: '' };
            }
            ws[cellAddr].s = cellStyle;
          }
        }
      }

      /* ── Anchos de columna ── */
      var colInfo = [];
      for (var ci = 0; ci < numCols; ci++) {
        var w = STATE.colWidths[ci];
        colInfo.push(w ? { wch: Math.round(w / 7) } : { wch: 14 });
      }
      ws['!cols'] = colInfo;

      XLSX.utils.book_append_sheet(wb, ws, STATE.activeSheet || 'Hoja1');

      var fname = (STATE.fileName.replace(/\.[^.]+$/, '') || 'documento') + '.xlsx';

      /* writeFile con bookSST:true y cellStyles:true para preservar estilos */
      XLSX.writeFile(wb, fname, { bookSST: true, cellStyles: true });

      STATE.dirty = false;
      updateUnsaved();
      showToast('Archivo guardado con estilos: ' + fname, 'success');
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
        ? '"' + s.replace(/"/g, '""') + '"'
        : s;
    }).join(',');
  }).join('\n');
  var blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' });
  var url  = URL.createObjectURL(blob);
  var a    = document.createElement('a');
  a.href     = url;
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
  var tbody = $('tableBody'); if (!tbody) return;

  tbody.querySelectorAll('.cell-display').forEach(function(span) {
    var td = span.closest('td'); if (!td) return;
    var r  = parseInt(td.getAttribute('data-row'));
    var c  = parseInt(td.getAttribute('data-col'));
    if (isNaN(r) || isNaN(c)) return;
    var rawSearch = String((STATE.data[r] ? STATE.data[r][c] : '') || '');
    if (rawSearch.toLowerCase().indexOf(lower) !== -1) {
      span.classList.add('search-match');
      _searchMatches.push({ span: span, r: r, c: c });
    }
  });

  $('searchCountBadge').textContent = _searchMatches.length ? String(_searchMatches.length) : '0';
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

function showFindReplace() { $('findReplacePanel').classList.remove('hidden'); $('frpFind').focus(); }
function hideFindReplace()  { $('findReplacePanel').classList.add('hidden'); }

function frpSearch() {
  var q       = $('frpFind').value;
  var caseSen = $('frpCaseSensitive').checked;
  var whole   = $('frpWholeCell').checked;
  _frpMatches = []; _frpIdx = -1;
  clearSearchHighlights();
  if (!q) { $('frpInfo').textContent = ''; $('frpInfo').className = 'frp-info'; return; }

  STATE.data.forEach(function(row, r) {
    row.forEach(function(val, c) {
      var v = String(val), s = q;
      if (!caseSen) { v = v.toLowerCase(); s = s.toLowerCase(); }
      var match = whole ? v === s : v.indexOf(s) !== -1;
      if (match) _frpMatches.push({ r: r, c: c });
    });
  });

  var info = $('frpInfo');
  if (_frpMatches.length) {
    info.textContent = _frpMatches.length + ' coincidencia(s)';
    info.className   = 'frp-info found';
    frpGoTo(0);
  } else {
    info.textContent = 'No encontrado';
    info.className   = 'frp-info none';
  }
}

function frpGoTo(idx) {
  if (!_frpMatches.length) return;
  _frpIdx = ((idx % _frpMatches.length) + _frpMatches.length) % _frpMatches.length;
  var match = _frpMatches[_frpIdx];
  activateCell(match.r, match.c);
  var span2 = getCellSpan(match.r, match.c);
  if (span2) {
    clearSearchHighlights();
    span2.classList.add('search-current');
  }
  $('frpInfo').textContent = (_frpIdx + 1) + ' / ' + _frpMatches.length;
}

function initFindReplace() {
  var frpClose         = $('frpClose');
  var frpFind          = $('frpFind');
  var frpFindNext      = $('frpFindNext');
  var frpFindPrev      = $('frpFindPrev');
  var frpCaseSensitive = $('frpCaseSensitive');
  var frpWholeCell     = $('frpWholeCell');
  var frpReplaceOne    = $('frpReplaceOne');
  var frpReplaceAll    = $('frpReplaceAll');

  if (frpClose)         frpClose.addEventListener('click',   hideFindReplace);
  if (frpFind)          frpFind.addEventListener('input',    frpSearch);
  if (frpFindNext)      frpFindNext.addEventListener('click', function() { frpGoTo(_frpIdx + 1); });
  if (frpFindPrev)      frpFindPrev.addEventListener('click', function() { frpGoTo(_frpIdx - 1); });
  if (frpCaseSensitive) frpCaseSensitive.addEventListener('change', frpSearch);
  if (frpWholeCell)     frpWholeCell.addEventListener('change',     frpSearch);

  if (frpReplaceOne) {
    frpReplaceOne.addEventListener('click', function() {
      if (_frpIdx < 0 || !_frpMatches.length) return frpSearch();
      var match2   = _frpMatches[_frpIdx];
      var q2       = $('frpFind').value;
      var rep2     = $('frpReplace').value;
      var caseSen2 = $('frpCaseSensitive').checked;
      var whole2   = $('frpWholeCell').checked;
      pushUndo();
      if (whole2) {
        STATE.data[match2.r][match2.c] = rep2;
      } else {
        var flags2 = caseSen2 ? 'g' : 'gi';
        var esc2   = q2.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        STATE.data[match2.r][match2.c] = String(STATE.data[match2.r][match2.c])
          .replace(new RegExp(esc2, flags2), rep2);
      }
      renderTable(); markDirty(); frpSearch();
      showToast('Reemplazado', 'success');
    });
  }

  if (frpReplaceAll) {
    frpReplaceAll.addEventListener('click', function() {
      var q3       = $('frpFind').value;
      var rep3     = $('frpReplace').value;
      var caseSen3 = $('frpCaseSensitive').checked;
      var whole3   = $('frpWholeCell').checked;
      if (!q3) return showToast('Escribe un término a buscar', '');
      pushUndo();
      var count2 = 0;
      var flags3 = caseSen3 ? 'g' : 'gi';
      var esc3   = q3.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      var regex3 = new RegExp(esc3, flags3);

      STATE.data.forEach(function(row, r) {
        row.forEach(function(val, c) {
          var v = String(val), s = q3;
          if (!caseSen3) { v = v.toLowerCase(); s = s.toLowerCase(); }
          var matches3 = whole3 ? v === s : v.indexOf(s) !== -1;
          if (matches3) {
            STATE.data[r][c] = whole3
              ? rep3
              : String(STATE.data[r][c]).replace(regex3, rep3);
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
   v7.8: añadida opción "Eliminar columna"
================================================ */
function initContextMenu() {
  document.addEventListener('click', function(e) {
    if (!e.target.closest('#contextMenu')) hideContextMenu();
  });
}

function showContextMenu(e, r, c) {
  activateCell(r, c);

  var ctx = $('contextMenu');
  ctx.innerHTML = '';

  var menuItems = [
    { label: 'Cortar',             icon: '&#9986;',   action: function() { copyCells(true); } },
    { label: 'Copiar',             icon: '&#128203;', action: function() { copyCells(false); } },
    { label: 'Pegar',              icon: '&#128204;', action: function() { pasteCells(false); } },
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
          else         nw[ki]     = STATE.colWidths[k];
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
          else        nw2[ki]     = STATE.colWidths[k];
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
          if (ki < c)      nw3[ki]     = STATE.colWidths[k];
          else if (ki > c) nw3[ki - 1] = STATE.colWidths[k];
        });
        STATE.colWidths   = nw3;
        STATE.activeCellC = Math.max(0, c - 1);
        STATE.selectedCells.clear();
        renderTable(); updateStats(); markDirty();
        showToast('Columna eliminada', 'info');
      }
    },
    { sep: true },
    { label: 'Agregar nota',   icon: '&#128221;', action: function() { addNote(r, c); } },
    { label: 'Limpiar celdas', icon: '&#10006;',  cls: 'danger', action: function() { clearActiveCells(); } },
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
    var cw = ctx.offsetWidth  || 200;
    var ch = ctx.offsetHeight || 300;
    var x  = Math.min(e.clientX, window.innerWidth  - cw - 10);
    var y  = Math.min(e.clientY, window.innerHeight - ch - 10);
    ctx.style.left = Math.max(4, x) + 'px';
    ctx.style.top  = Math.max(4, y) + 'px';
  });
}

function hideContextMenu() {
  var ctx = $('contextMenu');
  if (ctx) ctx.classList.add('hidden');
}

/* ================================================
   FILTER DROPDOWN
================================================ */
var _filterCol = -1;

function initFilterDropdown() {
  $('filterApply').addEventListener('click', applyColumnFilter);
  $('filterClear').addEventListener('click', function() {
    delete STATE.activeFilters[_filterCol];
    hideFilterDropdown();
    renderTable();
    var th = $('tableHead').querySelector('th[data-col="' + _filterCol + '"]');
    if (th) {
      var fi = th.querySelector('.filter-icon');
      if (fi) fi.classList.remove('active');
    }
    showToast('Filtro eliminado', 'info');
  });
  $('filterSearch').addEventListener('input', function() {
    var q = $('filterSearch').value.toLowerCase();
    $('filterOptions').querySelectorAll('.filter-option').forEach(function(opt) {
      var txt = opt.querySelector('.filter-option-text');
      opt.style.display = txt && txt.textContent.toLowerCase().indexOf(q) !== -1 ? '' : 'none';
    });
  });
}

function openFilterDropdown(c, iconEl) {
  _filterCol = c;
  var dd   = $('filterDropdown');
  var vals = [];
  var seen = {};
  STATE.data.forEach(function(row) {
    var v = String(row[c] || '');
    if (!seen[v]) { seen[v] = true; vals.push(v); }
  });
  vals.sort();
  var cur = STATE.activeFilters[c] || vals;

  $('filterSearch').value = '';
  var opts = $('filterOptions');
  opts.innerHTML = '';
  vals.forEach(function(v) {
    var div = document.createElement('div');
    div.className = 'filter-option';
    var cb  = document.createElement('input');
    cb.type    = 'checkbox';
    cb.value   = v;
    cb.checked = cur.indexOf(v) !== -1;
    var txt = document.createElement('span');
    txt.className   = 'filter-option-text';
    txt.textContent = v || '(vacío)';
    div.appendChild(cb);
    div.appendChild(txt);
    opts.appendChild(div);
  });

  var rect = iconEl.getBoundingClientRect();
  dd.style.left = Math.min(rect.left, window.innerWidth - 290) + 'px';
  dd.style.top  = (rect.bottom + 4) + 'px';
  dd.classList.remove('hidden');
}

function applyColumnFilter() {
  var checkboxes = $('filterOptions').querySelectorAll('input[type="checkbox"]:checked');
  var vals = Array.from(checkboxes).map(function(cb) { return cb.value; });
  if (!vals.length) {
    showToast('Selecciona al menos un valor para filtrar', '');
    return;
  }
  STATE.activeFilters[_filterCol] = vals;
  hideFilterDropdown();
  renderTable();
  showToast('Filtro aplicado', 'success');
}

function hideFilterDropdown() { $('filterDropdown').classList.add('hidden'); }

/* ================================================
   COLUMN RESIZE
================================================ */
function initColResize(th, colIdx, handle) {
  handle.addEventListener('mousedown', function(e) {
    e.preventDefault();
    e.stopPropagation();
    var startX = e.clientX;
    var startW = STATE.colWidths[colIdx] !== undefined
      ? STATE.colWidths[colIdx]
      : (th.offsetWidth || 90);

    document.body.style.cursor     = 'col-resize';
    document.body.style.userSelect = 'none';

    var onMove = function(ev) {
      var delta = ev.clientX - startX;
      var w = Math.max(20, startW + delta);
      STATE.colWidths[colIdx] = w;
      th.style.minWidth = w + 'px';
      th.style.width    = w + 'px';
      th.style.maxWidth = w + 'px';
      $('tableBody').querySelectorAll('td[data-col="' + colIdx + '"]').forEach(function(td) {
        td.style.width    = w + 'px';
        td.style.minWidth = w + 'px';
        td.style.maxWidth = w + 'px';
      });
    };
    var onUp = function() {
      document.body.style.cursor     = '';
      document.body.style.userSelect = '';
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup',   onUp);
    };
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup',   onUp);
  });
}

/* ================================================
   ROW RESIZE
================================================ */
function initRowResize(tr, rowIdx, handle) {
  handle.addEventListener('mousedown', function(e) {
    e.preventDefault(); e.stopPropagation();
    var startY = e.clientY;
    var startH = STATE.rowHeights[rowIdx] || tr.offsetHeight || 28;
    document.body.style.cursor = 'row-resize';

    var onMove = function(ev) {
      var h = Math.max(18, startH + ev.clientY - startY);
      tr.style.height          = h + 'px';
      STATE.rowHeights[rowIdx] = h;
    };
    var onUp = function() {
      document.body.style.cursor = '';
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
    };
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup',  onUp);
  });
}

/* ================================================
   FILL HANDLE (drag to fill)
================================================ */
function initFillHandle() {
  var scroll = $('tableScroll');
  scroll.addEventListener('mouseup', function() {
    if (STATE.isFillDragging && STATE.fillStart && STATE.fillEnd) {
      executeFillDown();
    }
    var tbody = $('tableBody');
    if (tbody) {
      tbody.querySelectorAll('.td-fill-preview').forEach(function(el) {
        el.classList.remove('td-fill-preview');
        el.style.outline = '';
      });
    }
    STATE.isFillDragging = false;
    STATE.fillStart      = null;
    STATE.fillEnd        = null;
    STATE.isDragging     = false;
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
    STATE.isDragging     = false;
    STATE.isFillDragging = false;
  });
  document.addEventListener('mouseleave', function() {
    STATE.isDragging     = false;
    STATE.isFillDragging = false;
  });
}

/* ================================================
   TOOLTIP
================================================ */
function initTooltip() {
  var tooltip = $('cellTooltip');
  if (!tooltip) return;
  var _ttTimer;

  document.addEventListener('mouseover', function(e) {
    var td = e.target.closest('#tableBody td[data-col]');
    if (!td) return;
    var r = parseInt(td.getAttribute('data-row'));
    var c = parseInt(td.getAttribute('data-col'));
    if (isNaN(r) || isNaN(c) || r < 0 || c < 0) return;
    var raw  = (STATE.data[r] ? STATE.data[r][c] : '') || '';
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
      tooltip.style.left = Math.min(e.clientX + 12, window.innerWidth  - 250) + 'px';
      tooltip.style.top  = Math.min(e.clientY + 18, window.innerHeight - 90)  + 'px';
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
var _noteDialog   = null;
var _noteCallback = null;

function initNoteDialog() {
  var dlg = document.createElement('div');
  dlg.id  = 'noteDialog';
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
    _noteCallback     = null;
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
    if (e.key === 'Escape') {
      dlg.style.display = 'none';
      _noteCallback = null;
    }
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
    if (note.trim()) {
      STATE.cellNotes[r + ',' + c] = note;
    } else {
      delete STATE.cellNotes[r + ',' + c];
    }
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
    dot.className  = 'note-indicator';
    dot.style.cssText = 'position:absolute;top:2px;right:2px;width:5px;height:5px;' +
      'border-radius:50%;background:var(--formula);pointer-events:none;z-index:2;';
    td.appendChild(dot);
  }
}

/* ================================================
   PRINT
================================================ */
function printPreview() {
  window.print();
}

/* ================================================
   RENAME FILE
================================================ */
function startRename() {
  var inp = $('fileNameInput');
  var lbl = $('fileName');
  inp.value = STATE.fileName;
  lbl.style.display = 'none';
  inp.style.display = '';
  $('btnRename').style.display = 'none';
  inp.focus(); inp.select();
}
function commitRename() {
  var inp = $('fileNameInput');
  var val = inp.value.trim();
  if (val) STATE.fileName = val;
  $('fileName').textContent    = STATE.fileName;
  $('fileName').style.display  = '';
  inp.style.display            = 'none';
  $('btnRename').style.display = '';
  markDirty();
}
function cancelRename() {
  $('fileNameInput').style.display = 'none';
  $('fileName').style.display      = '';
  $('btnRename').style.display     = '';
}
