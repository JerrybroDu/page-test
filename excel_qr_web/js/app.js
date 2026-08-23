/* ============================================================
 * 二维码批量生成器 · 网页版（Excel_Qr_Web）
 *
 * 完全在浏览器本地运行：选择 Excel → 设置参数 → 指定列文本生成
 * 二维码并新增一列插入 → 下载新文件。文件不会上传到任何服务器。
 *
 * 依赖（lib/ 目录，均已本地化，可离线使用）：
 *   - exceljs.min.js   读取 / 写入 .xlsx .xlsm，并支持在单元格插入图片
 *   - qrcode.min.js    二维码矩阵生成（支持 UTF-8 / 中文）
 *   - xlsx.full.min.js 仅用于读取旧版 .xls（BIFF）文件
 * ============================================================ */

(function () {
  'use strict';

  // ── 常量（与桌面版保持一致） ─────────────────────────────
  var SIZE_PRESETS = { '96': '小', '128': '中', '192': '大' };
  var LEVEL_MAP = { L: '低', M: '标准', Q: '高', H: '最高' };
  var BORDER_MAP = { '2': '紧凑', '4': '标准', '6': '宽松' };
  // 各纠错级别下字节模式最大容量（QR 规范 40 版本上限）
  var QR_MAX_BYTES = { L: 2953, M: 2331, Q: 1663, H: 1273 };
  var MAX_FILE_MB = 50;
  var PREVIEW_ROWS = 20;
  var PREVIEW_SCAN_LIMIT = 50000; // 预览扫描行数上限，避免超大表格卡死

  // 让 qrcode-generator 使用 UTF-8 字节编码（支持中文等）
  if (qrcode.stringToBytesFuncs && qrcode.stringToBytesFuncs['UTF-8']) {
    qrcode.stringToBytes = qrcode.stringToBytesFuncs['UTF-8'];
  }

  // ── 全局状态 ─────────────────────────────────────────────
  var state = {
    wb: null,           // ExcelJS.Workbook
    sourceBuffer: null, // 源文件原始字节（用于每次生成前重新加载，避免重复插入）
    fileName: '',       // 源文件名
    fileType: '',       // xlsx / xlsm / xls
    sheets: [],         // [{ name, headers, ncols, nrows, ws }]
    selSheetIdx: 0,
    headerMode: true,   // true=跳过首行（有表头）
    selCol: 1,
    sizePreset: '128',
    autoFit: true,
    // 高级设置
    customSize: false,
    customVal: 128,
    level: 'M',
    border: 4,
    customStart: false,
    startVal: 2,
    generating: false,
    lastImgSize: 0,
    pendingDownload: null, // 生成完成待确认下载的文件 { name, blob, count, skipped }
    saving: false,         // 正在保存/下载中（弹框进入转圈状态，禁止关闭）
  };

  // ── DOM 引用 ─────────────────────────────────────────────
  var $ = function (id) { return document.getElementById(id); };
  var el = {
    dropzone: $('dropzone'),
    fileInput: $('fileInput'),
    fileInfo: $('fileInfo'),
    fileName: $('fileName'),
    fileDetail: $('fileDetail'),
    btnReselect: $('btnReselect'),
    selSheet: $('selSheet'),
    selCol: $('selCol'),
    chkFit: $('chkFit'),
    btnAdvanced: $('btnAdvanced'),
    btnGenerate: $('btnGenerate'),
    btnCustomTop: $('btnCustomTop'),
    btnTheme: $('btnTheme'),
    themePanel: $('themePanel'),
    progressBar: $('progressBar'),
    progressText: $('progressText'),
    log: $('log'),
    previewHint: $('previewHint'),
    previewBody: $('previewBody'),
    // 高级设置
    chkCustomSize: $('chkCustomSize'),
    inpCustomSize: $('inpCustomSize'),
    selLevel: $('selLevel'),
    selBorder: $('selBorder'),
    chkCustomStart: $('chkCustomStart'),
    inpCustomStart: $('inpCustomStart'),
    // 自定义二维码
    txtQrContent: $('txtQrContent'),
    qrPreviewCanvas: $('qrPreviewCanvas'),
    qrPreviewEmpty: $('qrPreviewEmpty'),
    btnDownloadQr: $('btnDownloadQr'),
    // 下载确认
    dlModal: $('dlModal'),
    dlFileName: $('dlFileName'),
    dlLocation: $('dlLocation'),
    dlStats: $('dlStats'),
    dlHint: $('dlHint'),
    btnDlOk: $('btnDlOk'),
    btnDlCancel: $('btnDlCancel'),
    dlClose: $('dlClose'),
    dlLoading: $('dlLoading'),
  };

  // ══════════════════════════════════════════════════════════
  // 工具函数
  // ══════════════════════════════════════════════════════════

  function getExt(name) {
    var m = /\.([^.]+)$/.exec(name || '');
    return m ? m[1].toLowerCase() : '';
  }

  function pad2(n) { return String(n).padStart(2, '0'); }

  // Date → "YYYY-MM-DD HH:MM:SS"（与桌面版 xlrd 转换一致）
  function formatDate(d) {
    return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate()) +
      ' ' + pad2(d.getHours()) + ':' + pad2(d.getMinutes()) + ':' + pad2(d.getSeconds());
  }

  // 判断是否日期（跨 iframe / 跨 realm 也可靠）
  function isDateLike(v) {
    return Object.prototype.toString.call(v) === '[object Date]';
  }

  // 单元格值 → 纯文本（整数去掉小数尾巴由 String() 天然完成）
  function valueToText(value) {
    if (value === null || value === undefined) return '';
    var t = typeof value;
    if (t === 'string') return value;
    if (t === 'number') return String(value);
    if (t === 'boolean') return value ? 'TRUE' : 'FALSE';
    if (isDateLike(value)) return formatDate(value);
    if (t === 'object') {
      if (value.error !== undefined) return String(value.error);
      if (value.richText) return value.richText.map(function (r) { return r.text; }).join('');
      if (value.hyperlink !== undefined) return value.text !== undefined ? String(value.text) : String(value.hyperlink);
      if (value.result !== undefined) return valueToText(value.result);
      if (value.formula !== undefined) return String(value.formula);
      if (value.sharedFormula !== undefined) return valueToText(value.sharedFormula);
    }
    return String(value);
  }

  function cellText(cell) {
    return valueToText(cell && cell.value).trim();
  }

  // 内容字节数（UTF-8，用于判断是否超出二维码容量）
  function byteLength(str) {
    if (typeof TextEncoder === 'function') return new TextEncoder().encode(str).length;
    var n = 0;
    for (var i = 0; i < str.length; i++) {
      var c = str.charCodeAt(i);
      n += c < 0x80 ? 1 : c < 0x800 ? 2 : c < 0x10000 ? 3 : 4;
    }
    return n;
  }

  function yieldToUI() {
    return new Promise(function (r) { setTimeout(r, 0); });
  }

  // ── 日志 ───────────────────────────────────────────────
  function logLine(msg, cls) {
    var div = document.createElement('div');
    div.className = 'log-line' + (cls ? ' log-' + cls : '');
    div.textContent = msg;
    el.log.appendChild(div);
    if (el.log.children.length > 500) el.log.removeChild(el.log.firstChild);
    el.log.scrollTop = el.log.scrollHeight;
  }

  function log(msg, cls) {
    var t = new Date();
    var stamp = pad2(t.getHours()) + ':' + pad2(t.getMinutes()) + ':' + pad2(t.getSeconds());
    logLine('[' + stamp + '] ' + msg, cls);
  }

  function clearLog() {
    el.log.innerHTML = '';
  }

  // ── 进度条 ─────────────────────────────────────────────
  function setProgress(done, total, text) {
    var pct = total ? Math.round(done / total * 100) : 0;
    el.progressBar.style.setProperty('--progress', pct + '%');
    el.progressBar.setAttribute('aria-valuenow', String(pct));
    if (text) el.progressText.textContent = text;
  }

  function resetProgress() {
    setProgress(0, 1, '就绪');
  }

  // ── Toast 轻提示 ───────────────────────────────────────
  function toast(msg, type) {
    var old = document.querySelector('.toast');
    if (old) old.remove();
    var t = document.createElement('div');
    t.className = 'toast' + (type ? ' toast-' + type : '');
    t.textContent = msg;
    document.body.appendChild(t);
    setTimeout(function () { t.classList.add('toast-out'); setTimeout(function () { t.remove(); }, 300); }, 3200);
  }

  // ── 悬浮提示（ⓘ 图标） ────────────────────────────────
  var tooltipEl = $('tooltip');
  document.querySelectorAll('.help').forEach(function (el2) {
    el2.addEventListener('mouseenter', function () {
      var tip = el2.getAttribute('data-tip') || '';
      if (!tip) return;
      tooltipEl.textContent = tip;
      tooltipEl.classList.remove('hidden');
      var r = el2.getBoundingClientRect();
      var x = r.left + 20, y = r.top + 16;
      var tw = tooltipEl.offsetWidth, th = tooltipEl.offsetHeight;
      if (x + tw > window.innerWidth - 8) x = Math.max(8, r.left - tw - 12);
      if (y + th > window.innerHeight - 8) y = Math.max(8, r.top - th - 8);
      tooltipEl.style.left = x + 'px';
      tooltipEl.style.top = y + 'px';
    });
    el2.addEventListener('mouseleave', function () { tooltipEl.classList.add('hidden'); });
  });
  window.addEventListener('scroll', function () { tooltipEl.classList.add('hidden'); }, true);

  // ══════════════════════════════════════════════════════════
  // 文件选择
  // ══════════════════════════════════════════════════════════

  function isSupported(fileName) {
    return ['xlsx', 'xlsm', 'xls'].indexOf(getExt(fileName)) !== -1;
  }

  function openFileDialog() { el.fileInput.click(); }

  function handleDrop(e) {
    e.preventDefault();
    el.dropzone.classList.remove('dragover');
    var files = e.dataTransfer && e.dataTransfer.files;
    if (files && files.length) loadFile(files[0]);
  }

  el.dropzone.addEventListener('click', function () {
    if (!state.generating) openFileDialog();
  });
  el.dropzone.addEventListener('keydown', function (e) {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); openFileDialog(); }
  });
  el.dropzone.addEventListener('dragover', function (e) {
    e.preventDefault();
    el.dropzone.classList.add('dragover');
  });
  el.dropzone.addEventListener('dragleave', function () { el.dropzone.classList.remove('dragover'); });
  el.dropzone.addEventListener('drop', handleDrop);
  el.fileInput.addEventListener('change', function () {
    if (el.fileInput.files && el.fileInput.files.length) loadFile(el.fileInput.files[0]);
    el.fileInput.value = '';
  });
  el.btnReselect.addEventListener('click', function () { openFileDialog(); });

  // ── 读取文件 → ExcelJS 工作簿 ─────────────────────────
  async function loadFile(file) {
    if (!isSupported(file.name)) {
      toast('仅支持 .xlsx / .xlsm / .xls 格式', 'error');
      return;
    }
    if (file.size > MAX_FILE_MB * 1024 * 1024) {
      toast('文件超过 ' + MAX_FILE_MB + 'MB，请压缩后再试', 'error');
      return;
    }
    if (state.generating) return;
    clearLog();
    resetProgress();
    log('正在读取：' + file.name + ' …');
    try {
      var buf = await file.arrayBuffer();
      var wb;
      var ext = getExt(file.name);
      if (ext === 'xls') {
        wb = xlsToExcelJS(buf);
        log('旧版 .xls 已转换读取，结果将另存为 .xlsx', 'ok');
      } else {
        wb = new ExcelJS.Workbook();
        await wb.xlsx.load(buf);
      }
      state.wb = wb;
      state.sourceBuffer = buf;
      state.fileName = file.name;
      state.fileType = ext;
      state.sheets = buildSheetMeta(wb);
      if (!state.sheets.length) {
        throw new Error('该 Excel 文件没有任何工作表');
      }
      state.selSheetIdx = 0;
      state.selCol = 1;
      if (ext === 'xlsm') {
        log('提示：.xlsm 中的宏代码在另存为 .xlsx 后不会保留', '');
      }
      log('已读取 ' + state.sheets.length + ' 个工作表：' + state.sheets.map(function (s) { return s.name; }).join('、'), 'ok');
      showFileInfo();
      populateSheetSelect();
      el.btnGenerate.disabled = false;
      refreshControls();
    } catch (err) {
      log('读取 Excel 失败：' + (err && err.message ? err.message : err), 'error');
      toast('读取 Excel 失败，请检查文件', 'error');
    }
  }

  // 旧版 .xls（BIFF）→ ExcelJS 工作簿（值全部保留，样式不保留）
  function xlsToExcelJS(arrayBuffer) {
    var data = new Uint8Array(arrayBuffer);
    var wbx = XLSX.read(data, { type: 'array', cellDates: true });
    var wb = new ExcelJS.Workbook();
    wbx.SheetNames.forEach(function (name, i) {
      var wsName = (name || 'Sheet' + (i + 1)).slice(0, 31);
      var ws;
      try { ws = wb.addWorksheet(wsName); }
      catch (e) { ws = wb.addWorksheet('Sheet' + (i + 1)); }
      var aoa = XLSX.utils.sheet_to_json(wbx.Sheets[name], { header: 1, raw: true, defval: null });
      for (var r = 0; r < aoa.length; r++) {
        var rowArr = aoa[r];
        if (!rowArr) continue;
        for (var c = 0; c < rowArr.length; c++) {
          var v = rowArr[c];
          if (v === null || v === undefined || v === '') continue;
          ws.getCell(r + 1, c + 1).value = v;
        }
      }
    });
    return wb;
  }

  // 解析每个工作表的基础信息（表头、行列数）
  function buildSheetMeta(wb) {
    return wb.worksheets.map(function (ws) {
      var ncols = ws.columnCount;
      var headers = [];
      for (var c = 1; c <= ncols; c++) {
        headers.push(valueToText(ws.getCell(1, c).value));
      }
      while (headers.length && !headers[headers.length - 1]) headers.pop();
      return { name: ws.name, headers: headers, ncols: ncols, nrows: ws.rowCount, ws: ws };
    });
  }

  function currentSheet() {
    return state.sheets[state.selSheetIdx] || state.sheets[0];
  }

  function currentWs() {
    var s = currentSheet();
    return s ? s.ws : null;
  }

  // ══════════════════════════════════════════════════════════
  // 参数控件
  // ══════════════════════════════════════════════════════════

  function showFileInfo() {
    var s = currentSheet();
    el.fileName.textContent = state.fileName;
    el.fileDetail.textContent =
      state.sheets.length + ' 个工作表' +
      (s ? ' · 当前：' + s.name + '（' + s.nrows + ' 行 × ' + s.ncols + ' 列）' : '') +
      ' · ' + (state.fileType === 'xls' ? '旧版 .xls（将转存为 .xlsx）' : state.fileType.toUpperCase());
    el.fileInfo.classList.remove('hidden');
    el.dropzone.classList.add('hidden');
  }

  function populateSheetSelect() {
    el.selSheet.innerHTML = '';
    state.sheets.forEach(function (s, i) {
      var opt = document.createElement('option');
      opt.value = String(i);
      opt.textContent = s.name;
      el.selSheet.appendChild(opt);
    });
    el.selSheet.selectedIndex = state.selSheetIdx;
    el.selSheet.disabled = false;
  }

  function refreshControls() {
    var info = currentSheet();
    if (!info) return;
    refreshColChoices();
    renderPreview();
  }

  // 刷新“二维码内容列”下拉框，尽量保持用户已选列
  function refreshColChoices() {
    var info = currentSheet();
    if (!info) return;
    var headerMode = state.headerMode;
    var choices = [];
    if (headerMode && info.headers.length) {
      info.headers.forEach(function (h, i) {
        if (h) choices.push({ label: '第' + (i + 1) + '列（' + h.slice(0, 12) + '）', col: i + 1 });
      });
    }
    if (!choices.length) {
      for (var i = 1; i <= Math.max(info.ncols, 1); i++) {
        choices.push({ label: '第' + i + '列', col: i });
      }
    }
    // 当前选择仍在可选项里就保留，否则取第一项
    var cur = state.selCol;
    var exists = choices.some(function (c) { return c.col === cur; });
    if (!exists) { cur = choices[0].col; state.selCol = cur; }

    el.selCol.innerHTML = '';
    choices.forEach(function (c) {
      var opt = document.createElement('option');
      opt.value = String(c.col);
      opt.textContent = c.label;
      opt.selected = c.col === cur;
      el.selCol.appendChild(opt);
    });
    el.selCol.disabled = false;
  }

  function getStartRow() {
    return state.customStart ? state.startVal : (state.headerMode ? 2 : 1);
  }

  function getSizePx() {
    if (state.customSize) return state.customVal;
    return parseInt(state.sizePreset, 10) || 128;
  }

  // ── 数据预览 ─────────────────────────────────────────
  function renderPreview() {
    var info = currentSheet();
    var tbody = el.previewBody;
    if (!info) {
      tbody.innerHTML = '<tr><td colspan="3" class="empty-tip">请先选择 Excel 文件</td></tr>';
      el.previewHint.textContent = '';
      return;
    }
    var ws = info.ws;
    var colText = state.selCol;
    var startRow = getStartRow();
    var level = state.level;
    var sizePx = getSizePx();

    // 扫描前若干行，收集前 PREVIEW_ROWS 条非空内容
    var rows = [];
    var scanned = 0;
    var limit = Math.min(ws.rowCount, PREVIEW_SCAN_LIMIT);
    for (var r = startRow; r <= limit; r++) {
      var text = cellText(ws.getCell(r, colText));
      scanned++;
      if (text) {
        rows.push({ r: r, text: text });
        if (rows.length >= PREVIEW_ROWS) break;
      }
    }

    el.previewHint.textContent =
      (rows.length ? '第 ' + startRow + ' 行起 · 预览前 ' + rows.length + ' 条' : '该列从第 ' + startRow + ' 行起没有内容') +
      (scanned >= PREVIEW_SCAN_LIMIT ? '（已扫描 ' + scanned + ' 行）' : '');

    if (!rows.length) {
      tbody.innerHTML = '<tr><td colspan="3" class="empty-tip">该列从第 ' + startRow + ' 行起没有找到内容</td></tr>';
      return;
    }

    var html = [];
    rows.forEach(function (item) {
      var previewQr = '';
      try {
        var canvas = makeQR(item.text, Math.min(sizePx, 96), level, state.border);
        var url = canvas.toDataURL('image/png');
        // 预览用小图，宽高按实际画布等比缩放（保留锐利边缘）
        var w = Math.max(1, Math.round(48 * canvas.width / canvas.height));
        previewQr = '<img src="' + url + '" width="' + w + '" height="48" alt="" loading="lazy">';
      } catch (e) {
        previewQr = '<span class="preview-more">过长</span>';
      }
      html.push(
        '<tr><td class="col-row">' + item.r + '</td>' +
        '<td><span class="cell-elide" title="' + escAttr(item.text) + '">' + escHtml(item.text) + '</span></td>' +
        '<td class="col-qr">' + previewQr + '</td></tr>'
      );
    });
    tbody.innerHTML = html.join('');
  }

  function escHtml(s) {
    return String(s).replace(/[&<>"']/g, function (m) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[m];
    });
  }
  function escAttr(s) { return escHtml(s); }

  // ══════════════════════════════════════════════════════════
  // 二维码生成
  // ══════════════════════════════════════════════════════════

  // 生成二维码 canvas，按目标边长整数倍渲染模块（边缘锐利、无插值模糊）
  function makeQR(text, sizePx, level, border) {
    var qr = qrcode(0, level);   // typeNumber 0 = 自动选择最小版本
    qr.addData(text);
    qr.make();
    var n = qr.getModuleCount();
    var total = n + border * 2;
    var box = Math.max(1, Math.floor(sizePx / total));
    var size = box * total;
    var canvas = document.createElement('canvas');
    canvas.width = size;
    canvas.height = size;
    var ctx = canvas.getContext('2d');
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, size, size);
    ctx.fillStyle = '#000000';
    for (var gy = 0; gy < total; gy++) {
      for (var gx = 0; gx < total; gx++) {
        var qy = gy - border, qx = gx - border;
        if (qy >= 0 && qy < n && qx >= 0 && qx < n && qr.isDark(qy, qx)) {
          ctx.fillRect(gx * box, gy * box, box, box);
        }
      }
    }
    return canvas;
  }

  function canvasToPngBytes(canvas) {
    var b64 = canvas.toDataURL('image/png').split(',')[1];
    var bin = atob(b64);
    var bytes = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return bytes;
  }

  // 预检查：内容是否超出所选纠错级别下的二维码最大容量
  function checkContentLength(text, level, row) {
    var len = byteLength(text);
    if (len > QR_MAX_BYTES[level]) {
      var loc = row ? '第 ' + row + ' 行「' + text.slice(0, 20) + '」：' : '内容';
      throw new Error(loc + '过长（' + len + ' 字节），超出二维码最大容量：' +
        '纠错级别 ' + (LEVEL_MAP[level] || level) + ' 下最多约 ' + QR_MAX_BYTES[level] +
        ' 个字母/数字、' + Math.floor(QR_MAX_BYTES[level] / 3) + ' 个汉字');
    }
  }

  // ══════════════════════════════════════════════════════════
  // 生成主流程
  // ══════════════════════════════════════════════════════════

  function outputName(srcName) {
    var stem = String(srcName).replace(/\.[^.]+$/, '');
    var d = new Date();
    var stamp = pad2(d.getFullYear()) + pad2(d.getMonth() + 1) + pad2(d.getDate());
    return stem + '_二维码_' + stamp + '.xlsx';
  }

  function downloadBlob(blob, name) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(function () { URL.revokeObjectURL(url); a.remove(); }, 1500);
  }

  // ── 下载确认：生成完成后弹框询问是否下载，并展示保存位置 ──
  function canSavePicker() {
    // Chrome / Edge 支持原生「另存为」对话框，可让用户看到并自选保存路径
    return typeof window.showSaveFilePicker === 'function';
  }
  function showDownloadConfirm(info) {
    state.pendingDownload = info;
    state.saving = false;
    el.dlLoading.classList.add('hidden');
    el.dlFileName.textContent = info.name;
    el.dlStats.textContent = '共生成 ' + info.count + ' 个二维码' +
      (info.skipped ? '（跳过空行 ' + info.skipped + ' 个）' : '');
    el.dlLocation.textContent = '点击「下载文件」按钮，用户自定义存放文件位置';
    el.dlHint.textContent = canSavePicker()
      ? '支持自选保存位置（Chrome / Edge）'
      : '原文件不受影响';
    openModal('dlModal');
  }
  function closeDownloadConfirm(cancelled) {
    if (state.saving) return; // 保存中不允许关闭
    if (cancelled) log('已取消下载，文件未保存（如需结果请再次生成）', '');
    state.pendingDownload = null;
    closeModal('dlModal');
  }
  // Chrome / Edge：弹出系统「另存为」对话框，由用户选定实际保存路径
  async function saveBlobViaPicker(blob, name) {
    var handle = await window.showSaveFilePicker({
      suggestedName: name,
      types: [{
        description: 'Excel 文件',
        accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] },
      }],
    });
    var writable = await handle.createWritable();
    await writable.write(blob);
    await writable.close();
  }

  // 每次生成前从源文件字节重新加载工作簿（与桌面版“每次从磁盘重读”行为一致，
  // 避免上一次生成的图片残留、重复插入）
  async function reloadWorkbookFromSource() {
    if (!state.sourceBuffer) return;
    if (state.fileType === 'xls') {
      state.wb = xlsToExcelJS(state.sourceBuffer);
    } else {
      state.wb = new ExcelJS.Workbook();
      await state.wb.xlsx.load(state.sourceBuffer);
    }
    state.sheets = buildSheetMeta(state.wb);
  }

  async function generate() {
    if (state.generating) return;
    var info = currentSheet();
    if (!info || !state.wb) {
      toast('请先选择 Excel 文件', 'error');
      return;
    }
    await reloadWorkbookFromSource();
    info = currentSheet();
    var ws = info.ws;
    var colText = state.selCol;
    var level = state.level;
    var border = state.border;
    var sizePx = getSizePx();
    var startRow = getStartRow();
    var autoFit = state.autoFit;

    // 参数校验
    if (!(sizePx >= 32 && sizePx <= 1000)) {
      toast('二维码边长应在 32~1000 像素之间', 'error');
      return;
    }
    if (startRow < 1) {
      toast('起始行必须 ≥ 1', 'error');
      return;
    }

    // 二维码列：有内容的最后一列 + 1（新增列，不覆盖原数据）
    var maxCol = 0;
    ws.eachRow(function (row) {
      var vals = row.values;
      for (var i = vals.length - 1; i >= 0; i--) {
        if (vals[i] !== undefined && vals[i] !== null) {
          if (i > maxCol) maxCol = i;
          break;
        }
      }
    });
    var colQr = maxCol + 1;
    if (colQr === colText) {
      toast('二维码列不能与内容列相同', 'error');
      return;
    }

    // 收集内容列数据（跳过空行）
    var rows = [];
    var skipped = 0;
    var lastRow = ws.rowCount;
    for (var r = startRow; r <= lastRow; r++) {
      var text = cellText(ws.getCell(r, colText));
      if (text) {
        rows.push({ r: r, text: text });
      } else {
        skipped++;
      }
      if (r % 5000 === 0) await yieldToUI(); // 超大表格时保持界面响应
    }
    if (!rows.length) {
      toast('工作表「' + ws.name + '」第 ' + colText + ' 列从第 ' + startRow + ' 行起没有找到任何文本，请检查内容列与起始行设置', 'error');
      log('没有找到任何文本，已中止', 'error');
      return;
    }

    // 预检查全部内容长度，提前报错（不破坏工作表）
    for (var i = 0; i < rows.length; i++) {
      checkContentLength(rows[i].text, level, rows[i].r);
    }

    // 若跳过表头且二维码列表头为空，补一个“二维码”表头
    if (startRow > 1) {
      var headCell = ws.getCell(1, colQr);
      if (!valueToText(headCell.value).trim()) {
        headCell.value = '二维码';
        headCell.alignment = { horizontal: 'center', vertical: 'center' };
      }
    }

    state.generating = true;
    el.btnGenerate.disabled = true;
    clearLog();
    log('开始生成：' + state.fileName);
    log('工作表：' + ws.name + ' | 内容列 第' + colText + '列 | 二维码自动新增在最后一列之后');
    log('参数：边长 ' + sizePx + 'px | 纠错 ' + (LEVEL_MAP[level] || level) + ' | 留白 ' + (BORDER_MAP[border] || border) +
      ' | 起始行 ' + startRow + ' | 自动调整 ' + (autoFit ? '开' : '关'));
    log('输出文件：' + outputName(state.fileName));

    var qrCache = new Map(); // text → {bytes, size}，相同内容复用同一张图片
    var LOG_LIMIT = 200;

    try {
      for (i = 0; i < rows.length; i++) {
        var item = rows[i];
        var entry = qrCache.get(item.text);
        if (!entry) {
          var canvas = makeQR(item.text, sizePx, level, border);
          entry = { bytes: canvasToPngBytes(canvas), size: canvas.width };
          qrCache.set(item.text, entry);
        }
        if (autoFit) {
          var targetH = entry.size * 0.75 + 4;
          var fitRow = ws.getRow(item.r);
          if (!fitRow.height || fitRow.height < targetH) fitRow.height = targetH;
          var fitCol = ws.getColumn(colQr);
          var targetW = entry.size / 7 + 2;
          if (!fitCol.width || fitCol.width < targetW) fitCol.width = targetW;
        }

        // 图片在单元格内居中：偏移 = (单元格像素宽高 − 图片宽高) / 2（与桌面版一致）
        var cellCol = ws.getColumn(colQr);
        var cellRow = ws.getRow(item.r);
        var colPx = (cellCol.width || 8.43) * 7 + 5; // 列宽字符 → 像素（近似）
        var rowPx = (cellRow.height || 15) / 0.75;   // 行高磅 → 像素
        var offX = Math.max(0, Math.floor((colPx - entry.size) / 2));
        var offY = Math.max(0, Math.floor((rowPx - entry.size) / 2));
        var imgId = state.wb.addImage({ buffer: entry.bytes, extension: 'png' });
        // 注意：ExcelJS 浏览器版的锚点模型用 nativeCol/nativeRow（0 起始）+
        // nativeColOff/nativeRowOff（EMU，1px=9525）；colOff/rowOff 键会被丢弃
        ws.addImage(imgId, {
          tl: {
            nativeCol: colQr - 1,
            nativeColOff: offX * 9525,
            nativeRow: item.r - 1,
            nativeRowOff: offY * 9525,
          },
          ext: { width: entry.size, height: entry.size },
        });

        setProgress(i + 1, rows.length, '生成 ' + (i + 1) + '/' + rows.length);
        if (i < LOG_LIMIT) {
          log('生成 ' + (i + 1) + '/' + rows.length + '：' + item.text);
        } else if (i === LOG_LIMIT) {
          log('… 共 ' + rows.length + ' 条，后续仅显示进度，不再逐条打印');
        }

        if (i % 20 === 19) await yieldToUI();
      }

      log('正在写文件并打包 …');
      await yieldToUI();
      var out = await state.wb.xlsx.writeBuffer();
      var name = outputName(state.fileName);
      var blob = new Blob([out], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });

      setProgress(rows.length, rows.length, '完成');
      log('[完成] 共生成 ' + rows.length + ' 个二维码（跳过空行 ' + skipped + ' 个）', 'ok');
      log('结果文件：' + name + '（等待确认下载）', 'ok');
      showDownloadConfirm({ name: name, blob: blob, count: rows.length, skipped: skipped });
      toast('生成完成，请确认是否下载', 'ok');
    } catch (err) {
      var msg = err && err.message ? err.message : String(err);
      log('[错误] ' + msg, 'error');
      setProgress(0, 1, '出错');
      toast('生成失败：' + msg, 'error');
    } finally {
      state.generating = false;
      el.btnGenerate.disabled = false;
    }
  }

  // ══════════════════════════════════════════════════════════
  // 控件事件
  // ══════════════════════════════════════════════════════════

  el.selSheet.addEventListener('change', function () {
    state.selSheetIdx = parseInt(el.selSheet.value, 10) || 0;
    showFileInfo();
    refreshControls();
  });

  document.querySelectorAll('#radioHeader input').forEach(function (r) {
    r.addEventListener('change', function () {
      state.headerMode = r.value === '1';
      refreshControls();
    });
  });

  document.querySelectorAll('#radioSize input').forEach(function (r) {
    r.addEventListener('change', function () { state.sizePreset = r.value; renderPreview(); });
  });

  el.selCol.addEventListener('change', function () {
    state.selCol = parseInt(el.selCol.value, 10) || 1;
    renderPreview();
  });

  el.chkFit.addEventListener('change', function () { state.autoFit = el.chkFit.checked; });

  el.btnGenerate.addEventListener('click', generate);

  // ── 高级设置对话框 ──────────────────────────────────
  el.btnAdvanced.addEventListener('click', function () { openModal('advModal'); });

  el.chkCustomSize.addEventListener('change', function () {
    state.customSize = el.chkCustomSize.checked;
    el.inpCustomSize.disabled = !state.customSize;
  });
  el.inpCustomSize.addEventListener('change', function () {
    var v = parseInt(el.inpCustomSize.value, 10);
    if (!(v >= 32 && v <= 1000)) { toast('边长应在 32~1000 之间', 'error'); v = 128; el.inpCustomSize.value = '128'; }
    state.customVal = v;
    renderPreview();
  });

  el.selLevel.addEventListener('change', function () {
    state.level = el.selLevel.value;
    renderPreview();
  });
  el.selBorder.addEventListener('change', function () {
    state.border = parseInt(el.selBorder.value, 10) || 4;
    renderPreview();
  });

  el.chkCustomStart.addEventListener('change', function () {
    state.customStart = el.chkCustomStart.checked;
    el.inpCustomStart.disabled = !state.customStart;
    renderPreview();
  });
  el.inpCustomStart.addEventListener('change', function () {
    var v = parseInt(el.inpCustomStart.value, 10);
    if (!(v >= 1)) { v = 2; el.inpCustomStart.value = '2'; }
    state.startVal = v;
    renderPreview();
  });

  // ── 模态框开关 ──────────────────────────────────────
  function openModal(id) {
    var m = $(id);
    m.classList.remove('hidden');
    document.body.classList.add('modal-open'); // 锁定背景滚动，等待用户处理完弹框
    var sel = m.querySelector('select, input, textarea');
    if (sel && sel.focus) setTimeout(function () { sel.focus(); }, 50);
  }
  function closeModal(id) {
    $(id).classList.add('hidden');
    if (!document.querySelector('.modal-mask:not(.hidden)')) {
      document.body.classList.remove('modal-open'); // 全部弹框关闭后再恢复滚动
    }
  }
  function closeAllModals() {
    document.querySelectorAll('.modal-mask:not(.hidden)').forEach(function (m) {
      closeModal(m.id);
    });
  }
  document.querySelectorAll('[data-close]').forEach(function (b) {
    b.addEventListener('click', function () { closeModal(b.getAttribute('data-close')); });
  });
  document.querySelectorAll('.modal-mask').forEach(function (m) {
    m.addEventListener('mousedown', function (e) {
      if (e.target === m && m.id !== 'dlModal') closeModal(m.id); // dlModal 单独处理（含保存中保护）
    });
  });
  document.addEventListener('keydown', function (e) {
    if (e.key !== 'Escape') return;
    // 下载确认框按 Esc 视为“取消下载”，清理待下载内容
    if (!el.dlModal.classList.contains('hidden')) { closeDownloadConfirm(true); return; }
    closeAllModals();
  });

  // ── 下载确认框按钮 ──────────────────────────────
  el.btnDlOk.addEventListener('click', async function () {
    var info = state.pendingDownload;
    if (!info || state.saving) return;
    state.saving = true;
    el.btnDlOk.disabled = true;
    el.btnDlCancel.disabled = true;
    el.dlLoading.classList.remove('hidden'); // 转圈：正在保存
    try {
      if (canSavePicker()) {
        // Chrome / Edge：系统「另存为」对话框，用户可看到并选择实际保存路径
        await saveBlobViaPicker(info.blob, info.name);
        await yieldToUI(); // 保证转圈至少渲染一帧再关闭
        log('文件已保存：' + info.name, 'ok');
      } else {
        downloadBlob(info.blob, info.name);
        await new Promise(function (r) { setTimeout(r, 700); }); // 转圈短暂反馈
        log('文件已下载：' + info.name, 'ok');
      }
      state.pendingDownload = null;
      el.dlLoading.classList.add('hidden');
      closeModal('dlModal');
      toast('下载成功：' + info.name, 'ok');
    } catch (err) {
      el.dlLoading.classList.add('hidden');
      if (err && err.name === 'AbortError') {
        log('已取消保存', '');
      } else {
        log('保存失败：' + (err && err.message ? err.message : err), 'error');
        toast('保存失败：' + (err && err.message ? err.message : err), 'error');
      }
    } finally {
      state.saving = false;
      el.btnDlOk.disabled = false;
      el.btnDlCancel.disabled = false;
    }
  });
  function cancelDownloadDialog() { closeDownloadConfirm(true); }
  el.btnDlCancel.addEventListener('click', cancelDownloadDialog);
  el.dlClose.addEventListener('click', cancelDownloadDialog);
  el.dlModal.addEventListener('mousedown', function (e) {
    if (e.target === el.dlModal) cancelDownloadDialog();
  });

  // ── 主题配色切换（右上角调色板按钮） ──────────────────────
  var THEMES = [
    { id: 'green',  name: '微信绿', color: '#07c160' },
    { id: 'blue',   name: '清爽蓝', color: '#33a5e8' },
    { id: 'orange', name: '活力橙', color: '#f59b18' },
    { id: 'purple', name: '科技紫', color: '#a26ff7' },
    { id: 'teal',   name: '翡翠青', color: '#17b8a6' },
  ];
  function themeColor(id) {
    for (var i = 0; i < THEMES.length; i++) if (THEMES[i].id === id) return THEMES[i].color;
    return '#07c160';
  }
  function applyTheme(id) {
    document.documentElement.setAttribute('data-theme', id);
    try { localStorage.setItem('qr_theme', id); } catch (e) {}
    var items = el.themePanel.querySelectorAll('.theme-item');
    for (var i = 0; i < items.length; i++) {
      items[i].classList.toggle('active', items[i].getAttribute('data-theme') === id);
    }
  }
  function closeThemePanel() {
    el.themePanel.classList.add('hidden');
    el.btnTheme.setAttribute('aria-expanded', 'false');
  }
  // 打开时把下拉面板固定在视口内（跟随按钮右下角，夹取不越界），
  // 避免手机窄屏 / 页面滚动 / 屏幕旋转时溢出屏幕
  function positionThemePanel() {
    var panel = el.themePanel;
    var btnRect = el.btnTheme.getBoundingClientRect();
    var vw = window.innerWidth, vh = window.innerHeight;
    var pw = panel.offsetWidth;
    var left = Math.max(8, btnRect.right - pw);
    left = Math.min(left, vw - pw - 8);
    var top = Math.min(btnRect.bottom + 8, Math.max(8, vh - panel.offsetHeight - 8));
    panel.style.position = 'fixed';
    panel.style.left = left + 'px';
    panel.style.right = 'auto';
    panel.style.top = top + 'px';
  }
  function initTheme() {
    var saved = 'green';
    try { saved = localStorage.getItem('qr_theme') || 'green'; } catch (e) {}
    if (themeColor(saved) === '#07c160' && saved !== 'green') saved = 'green'; // 无效值回退
    applyTheme(saved);

    el.btnTheme.addEventListener('click', function (ev) {
      ev.stopPropagation();
      var willOpen = el.themePanel.classList.contains('hidden');
      el.themePanel.classList.toggle('hidden', !willOpen);
      if (willOpen) positionThemePanel();
      el.btnTheme.setAttribute('aria-expanded', willOpen ? 'true' : 'false');
    });
    // 面板打开时，窗口缩放 / 滚动 / 旋转后重新夹取位置
    window.addEventListener('resize', function () {
      if (!el.themePanel.classList.contains('hidden')) positionThemePanel();
    });
    window.addEventListener('scroll', function () {
      if (!el.themePanel.classList.contains('hidden')) positionThemePanel();
    }, true);
    el.themePanel.querySelectorAll('.theme-item').forEach(function (item) {
      item.addEventListener('click', function (ev) {
        ev.stopPropagation();
        applyTheme(item.getAttribute('data-theme'));
        closeThemePanel();
      });
    });
    document.addEventListener('click', function () { closeThemePanel(); });
    document.addEventListener('keydown', function (e) {
      if (e.key === 'Escape') closeThemePanel();
    });
  }
  initTheme();

  // ── 自定义二维码对话框 ──────────────────────────────
  var qrTimer = null;
  function refreshCustomQR() {
    var content = el.txtQrContent.value.replace(/\s+$/g, '');
    var ctx = el.qrPreviewCanvas.getContext('2d');
    if (!content) {
      el.qrPreviewEmpty.classList.remove('hidden');
      ctx.clearRect(0, 0, el.qrPreviewCanvas.width, el.qrPreviewCanvas.height);
      el.btnDownloadQr.disabled = true;
      return;
    }
    try {
      checkContentLength(content, 'M');
      var canvas = makeQR(content, 240, 'M', 4);
      ctx.clearRect(0, 0, el.qrPreviewCanvas.width, el.qrPreviewCanvas.height);
      ctx.drawImage(canvas, 0, 0);
      el.qrPreviewEmpty.classList.add('hidden');
      el.btnDownloadQr.disabled = false;
    } catch (e) {
      ctx.clearRect(0, 0, el.qrPreviewCanvas.width, el.qrPreviewCanvas.height);
      el.qrPreviewEmpty.textContent = '生成失败：' + (e && e.message ? e.message : e);
      el.qrPreviewEmpty.classList.remove('hidden');
      el.btnDownloadQr.disabled = true;
    }
  }
  el.txtQrContent.addEventListener('input', function () {
    if (qrTimer) clearTimeout(qrTimer);
    qrTimer = setTimeout(refreshCustomQR, 400);
  });
  el.btnCustomTop.addEventListener('click', function () {
    el.txtQrContent.value = '';
    refreshCustomQR();
    openModal('qrModal');
  });
  el.btnDownloadQr.addEventListener('click', function () {
    var content = el.txtQrContent.value.replace(/\s+$/g, '');
    if (!content) return;
    try {
      var canvas = makeQR(content, 512, 'M', 4);
      var url = canvas.toDataURL('image/png');
      var a = document.createElement('a');
      a.href = url;
      a.download = '二维码.png';
      document.body.appendChild(a);
      a.click();
      a.remove();
      log('二维码图片已导出：二维码.png', 'ok');
    } catch (e) {
      toast('导出失败：' + (e && e.message ? e.message : e), 'error');
    }
  });

  // ── 初始化 ──────────────────────────────────────────
  document.addEventListener('DOMContentLoaded', function () {
    log('二维码批量生成器 · 网页版 已就绪。所有处理均在本地完成，文件不会上传。');
  });

  // 供测试脚本引用（Node 环境无 DOM 时可调用核心逻辑）
  window.__qrExports = {
    makeQR: makeQR,
    valueToText: valueToText,
    cellText: cellText,
    checkContentLength: checkContentLength,
    canvasToPngBytes: canvasToPngBytes,
    byteLength: byteLength,
    outputName: outputName,
    showDownloadConfirm: showDownloadConfirm,
    canSavePicker: canSavePicker,
  };
})();
