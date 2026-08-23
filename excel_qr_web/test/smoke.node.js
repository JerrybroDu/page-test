#!/usr/bin/env node
/**
 * 命令行冒烟测试（Node.js，可选开发工具）
 *
 * 验证网页版的核心流程：读取 Excel → 按列生成二维码 PNG → 新增一列并
 * 把二维码图片插入对应单元格 → 另存为 .xlsx。测试中使用纯 JS 灰度 PNG
 * 编码器替代浏览器的 canvas，二维码模块布局与 js/app.js 中的 makeQR 一致。
 *
 * 运行：node test/smoke.node.js
 * 输出：test/output_smoke.xlsx
 */
'use strict';

const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

const ROOT = path.join(__dirname, '..');
const LIB = path.join(ROOT, 'lib');
const ExcelJS = require(path.join(LIB, 'exceljs.min.js'));
const qrcode = require(path.join(LIB, 'qrcode.min.js'));
const { makeGrayscalePng } = require('./pngenc.node.js');

qrcode.stringToBytes = qrcode.stringToBytesFuncs['UTF-8'];

const QR_MAX_BYTES = { L: 2953, M: 2331, Q: 1663, H: 1273 };

// ── 与 js/app.js 的 makeQR 相同的坐标算法（灰度 PNG 版） ──
function makeQR(text, sizePx, level, border) {
  const qr = qrcode(0, level);
  qr.addData(text);
  qr.make();
  const n = qr.getModuleCount();
  const total = n + border * 2;
  const box = Math.max(1, Math.floor(sizePx / total));
  const size = box * total;
  const rows = [];
  for (let y = 0; y < size; y++) {
    const row = [];
    for (let x = 0; x < size; x++) {
      const gy = Math.floor(y / box), gx = Math.floor(x / box);
      const qy = gy - border, qx = gx - border;
      row.push(qy >= 0 && qy < n && qx >= 0 && qx < n && qr.isDark(qy, qx));
    }
    rows.push(row);
  }
  return makeGrayscalePng(rows, size);
}

// ── 与 js/app.js 的 valueToText 相同的文本规整 ──
function valueToText(v) {
  if (v === null || v === undefined) return '';
  const t = typeof v;
  if (t === 'string') return v;
  if (t === 'number') return String(v);
  if (t === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (v instanceof Date) {
    const p = (n) => String(n).padStart(2, '0');
    return v.getFullYear() + '-' + p(v.getMonth() + 1) + '-' + p(v.getDate()) +
      ' ' + p(v.getHours()) + ':' + p(v.getMinutes()) + ':' + p(v.getSeconds());
  }
  if (t === 'object') {
    if (v.error !== undefined) return String(v.error);
    if (v.richText) return v.richText.map((r) => r.text).join('');
    if (v.hyperlink !== undefined) return v.text !== undefined ? String(v.text) : String(v.hyperlink);
    if (v.result !== undefined) return valueToText(v.result);
    if (v.formula !== undefined) return String(v.formula);
  }
  return String(v);
}

function assert(cond, msg) {
  if (!cond) { console.error('✗ FAIL: ' + msg); process.exit(1); }
  console.log('  ✓ ' + msg);
}

(async () => {
  const src = path.join(ROOT, '测试数据.xlsx');
  assert(fs.existsSync(src), '找到测试数据 ' + path.basename(src));

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(fs.readFileSync(src));
  assert(wb.worksheets.length === 5, '读取 5 个工作表');

  const ws = wb.getWorksheet(1);
  const colText = 1;           // 姓名列
  const startRow = 2;          // 跳过表头
  const level = 'M', border = 4, sizePx = 128;

  // 1. 定位二维码列 = 有内容的最后一列 + 1
  let maxCol = 0;
  ws.eachRow((row) => {
    const vals = row.values;
    for (let i = vals.length - 1; i >= 0; i--) {
      if (vals[i] !== undefined && vals[i] !== null) { if (i > maxCol) maxCol = i; break; }
    }
  });
  const colQr = maxCol + 1;
  assert(colQr === 4, '二维码列定位到第 4 列（最后一列之后）');

  // 2. 收集数据
  const rows = [];
  for (let r = startRow; r <= ws.rowCount; r++) {
    const text = valueToText(ws.getCell(r, colText).value).trim();
    if (text) rows.push({ r, text });
  }
  assert(rows.length >= 5, '收集到内容行（>=' + 5 + '），实际 ' + rows.length);

  // 3. 表头补写
  ws.getCell(1, colQr).value = '二维码';
  ws.getCell(1, colQr).alignment = { horizontal: 'center', vertical: 'center' };

  // 4. 逐行生成并插入二维码（先自动调整，再计算居中偏移，与 js/app.js 一致）
  const qrCache = new Map();
  const expectAnchors = []; // 记录每张图的期望锚点（供回读校验居中）
  for (const item of rows) {
    const bytes = new TextEncoder().encode(item.text);
    assert(bytes.length <= QR_MAX_BYTES[level], '内容长度未超限（' + item.text.slice(0, 12) + '）');
    let entry = qrCache.get(item.text);
    if (!entry) {
      entry = { png: makeQR(item.text, sizePx, level, border), size: 0 };
      // 读取 PNG 宽高作为实际尺寸（与浏览器 canvas.width 一致）
      entry.size = entry.png.readUInt32BE(16);
      qrCache.set(item.text, entry);
    }
    // 自动调整行高列宽（与 js/app.js 相同公式）
    const targetH = entry.size * 0.75 + 4;
    const fitRow = ws.getRow(item.r);
    if (!fitRow.height || fitRow.height < targetH) fitRow.height = targetH;
    const fitCol = ws.getColumn(colQr);
    const targetW = entry.size / 7 + 2;
    if (!fitCol.width || fitCol.width < targetW) fitCol.width = targetW;
    // 图片在单元格内居中：偏移 = (单元格像素宽高 − 图片宽高) / 2
    const colPx = (ws.getColumn(colQr).width || 8.43) * 7 + 5; // 列宽字符 → 像素（近似）
    const rowPx = (ws.getRow(item.r).height || 15) / 0.75;     // 行高磅 → 像素
    const offX = Math.max(0, Math.floor((colPx - entry.size) / 2));
    const offY = Math.max(0, Math.floor((rowPx - entry.size) / 2));
    expectAnchors.push({ r: item.r, offX, offY });
    const imgId = wb.addImage({ buffer: entry.png, extension: 'png' });
    // 锚点模型用 nativeCol/nativeRow（0 起始）+ nativeColOff/nativeRowOff（EMU）
    ws.addImage(imgId, {
      tl: { nativeCol: colQr - 1, nativeColOff: offX * 9525, nativeRow: item.r - 1, nativeRowOff: offY * 9525 },
      ext: { width: entry.size, height: entry.size },
    });
  }

  // 5. 写出
  const out = Buffer.from(await wb.xlsx.writeBuffer());
  const outPath = path.join(__dirname, 'output_smoke.xlsx');
  fs.writeFileSync(outPath, out);
  assert(out.length > 0, '写出文件 ' + path.basename(outPath) + '（' + out.length + ' 字节）');

  // 6. 回读校验
  const wb2 = new ExcelJS.Workbook();
  await wb2.xlsx.load(out);
  assert(wb2.worksheets.length === 5, '回读后仍为 5 个工作表');
  const ws2 = wb2.getWorksheet(1);
  assert(ws2.getCell(1, 4).value === '二维码', '表头为「二维码」');
  const imgs = ws2.getImages();
  assert(imgs.length === rows.length, '插入二维码图片 ' + imgs.length + ' 张（= 内容行数 ' + rows.length + '）');
  const first = imgs[0].range.tl;
  assert(first.nativeCol === 3 && first.nativeRow === 1,
    '首张二维码锚定在 D2（nativeCol=' + first.nativeCol + ',nativeRow=' + first.nativeRow + '）');
  const exp = expectAnchors[0];
  assert(first.nativeColOff === exp.offX * 9525 && first.nativeRowOff === exp.offY * 9525,
    'D2 单元格内居中（offX=' + exp.offX + 'px, offY=' + exp.offY + 'px → nativeColOff=' + first.nativeColOff + ', nativeRowOff=' + first.nativeRowOff + ' EMU）');

  console.log('\n冒烟测试通过 ✔');
})().catch((e) => {
  console.error('冒烟测试异常：', e);
  process.exit(1);
});
