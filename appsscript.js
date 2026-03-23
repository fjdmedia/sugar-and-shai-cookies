function doPost(e) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Orders') || ss.getActiveSheet();

  if (sheet.getLastRow() === 0) {
    setupSheet(sheet);
  }

  const p = e.parameter;
  sheet.appendRow([
    new Date().toLocaleString(),
    p.name        || '',
    p.email       || '',
    p.contact     || '',
    p.pickup_date || '',
    p.flavours    || '',
    p.quantity    || '',
    p.payment     || '',
    p.notes       || ''
  ]);

  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 1, 1, 9)
    .setBackground(lastRow % 2 === 0 ? '#FDF6EE' : '#FFFFFF');

  return ContentService
    .createTextOutput(JSON.stringify({ result: 'success' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────────────────────────────────
//  DASHBOARD  — run manually any time to rebuild
// ─────────────────────────────────────────────────────────────────
function buildDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Find the orders sheet — try by name first, then fall back to any sheet with 9 columns of data
  let ordersSheet = ss.getSheetByName('Orders');
  if (!ordersSheet) {
    ss.getSheets().forEach(s => {
      if (!ordersSheet && s.getLastRow() > 0 && s.getLastColumn() >= 9) {
        ordersSheet = s;
      }
    });
  }
  if (!ordersSheet) { Logger.log('No orders sheet found.'); return; }

  // Get or create Dashboard tab
  let dash = ss.getSheetByName('Dashboard');
  if (!dash) {
    dash = ss.insertSheet('Dashboard');
    ss.moveActiveSheet(1);
  }

  // Full reset
  try { dash.getRange(1, 1, dash.getMaxRows(), dash.getMaxColumns()).breakApart(); } catch(e) {}
  dash.clear();
  dash.setTabColor('#C97B2A');

  // ── Read orders ──
  const lastRow = ordersSheet.getLastRow();
  const orders  = lastRow > 1
    ? ordersSheet.getRange(2, 1, lastRow - 1, 9).getValues()
    : [];

  const totalOrders  = orders.length;
  let   totalDozens  = 0;
  const paymentMap   = {};
  const flavourMap   = {};
  const monthlyMap   = {};

  orders.forEach(r => {
    const qty = parseFloat(r[6]);
    if (!isNaN(qty)) totalDozens += qty;

    const method = (r[7] || '').toString().trim() || 'Unknown';
    paymentMap[method] = (paymentMap[method] || 0) + 1;

    (r[5] || '').toString().split(',').forEach(f => {
      const name = f.trim();
      if (name && name !== 'None selected') {
        flavourMap[name] = (flavourMap[name] || 0) + 1;
      }
    });

    const d = new Date(r[0]);
    if (!isNaN(d)) {
      const key = Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMM yyyy');
      monthlyMap[key] = (monthlyMap[key] || 0) + 1;
    }
  });

  // ── Colours ──
  const OLIVE = '#4A5E2A';
  const AMBER = '#C97B2A';
  const MOCHA = '#5A3E2A';
  const CREAM = '#FDF6EE';
  const WHITE = '#FFFFFF';
  const BROWN = '#3D2B27';
  const LGOLD = '#F0E6D3';

  // ── Layout: 4 columns ──
  dash.setColumnWidth(1, 30);   // left margin
  dash.setColumnWidth(2, 240);  // label / name
  dash.setColumnWidth(3, 160);  // value / count
  dash.setColumnWidth(4, 30);   // spacer
  dash.setColumnWidth(5, 240);  // label / name  (right panel)
  dash.setColumnWidth(6, 160);  // value / count (right panel)
  dash.setColumnWidth(7, 30);   // right margin

  let row = 1;

  // ── Helpers ──
  function banner(text, bg, size) {
    size = size || 13;
    dash.getRange(row, 1, 1, 7).merge()
      .setValue(text)
      .setBackground(bg)
      .setFontColor(WHITE)
      .setFontWeight('bold')
      .setFontSize(size)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    dash.setRowHeight(row, size >= 14 ? 50 : 34);
    row++;
  }

  function spacer(h) {
    dash.setRowHeight(row, h || 12);
    row++;
  }

  // Left-panel section header (cols 2-3)
  function leftHeader(text, bg) {
    dash.getRange(row, 2, 1, 2).merge()
      .setValue(text)
      .setBackground(bg)
      .setFontColor(WHITE)
      .setFontWeight('bold')
      .setFontSize(10)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    dash.setRowHeight(row, 30);
  }

  // Right-panel section header (cols 5-6)
  function rightHeader(text, bg) {
    dash.getRange(row, 5, 1, 2).merge()
      .setValue(text)
      .setBackground(bg)
      .setFontColor(WHITE)
      .setFontWeight('bold')
      .setFontSize(10)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    dash.setRowHeight(row, 30);
  }

  function leftRow(label, value, bg) {
    dash.getRange(row, 2).setValue(label)
      .setBackground(bg).setFontColor(BROWN).setFontWeight('bold')
      .setVerticalAlignment('middle').setHorizontalAlignment('left');
    dash.getRange(row, 3).setValue(value)
      .setBackground(bg).setFontColor(BROWN)
      .setVerticalAlignment('middle').setHorizontalAlignment('left');
    dash.setRowHeight(row, 26);
  }

  function rightRow(label, value, bg) {
    dash.getRange(row, 5).setValue(label)
      .setBackground(bg).setFontColor(BROWN).setFontWeight('bold')
      .setVerticalAlignment('middle').setHorizontalAlignment('left');
    dash.getRange(row, 6).setValue(value)
      .setBackground(bg).setFontColor(BROWN)
      .setVerticalAlignment('middle').setHorizontalAlignment('left');
    dash.setRowHeight(row, 26);
  }

  // ══════════════════════════════════════
  //  TITLE
  // ══════════════════════════════════════
  banner('🍪  Sugar & Shai Cookies — Orders Dashboard', OLIVE, 15);
  spacer(8);

  // ══════════════════════════════════════
  //  ROW A: Summary (left) + Monthly (right)  — share a row block
  // ══════════════════════════════════════
  const flavourEntries = Object.entries(flavourMap).sort((a,b) => b[1]-a[1]);
  const paymentEntries = Object.entries(paymentMap).sort((a,b) => b[1]-a[1]);
  const monthlyEntries = Object.entries(monthlyMap);

  leftHeader('📊  Summary', OLIVE);
  rightHeader('📅  Orders by Month', OLIVE);
  row++;

  const summaryData = [
    ['Total Orders',         totalOrders],
    ['Total Dozens Ordered', totalDozens],
    ['Most Popular Flavour', topEntry(flavourMap) || '—'],
    ['Top Payment Method',   topEntry(paymentMap) || '—'],
  ];

  const maxSummaryMonthly = Math.max(summaryData.length, monthlyEntries.length, 1);
  for (let i = 0; i < maxSummaryMonthly; i++) {
    const bg = i % 2 === 0 ? CREAM : WHITE;
    if (summaryData[i]) leftRow(summaryData[i][0], summaryData[i][1], bg);
    if (monthlyEntries[i]) rightRow(monthlyEntries[i][0], monthlyEntries[i][1] + ' orders', bg);
    else if (!monthlyEntries.length && i === 0) rightRow('No data yet', '', CREAM);
    row++;
  }

  spacer(16);

  // ══════════════════════════════════════
  //  ROW B: Flavours (left) + Payment (right)
  // ══════════════════════════════════════
  leftHeader('🍫  Flavour Popularity', AMBER);
  rightHeader('💳  Payment Methods', MOCHA);
  row++;

  const maxFlavPay = Math.max(flavourEntries.length, paymentEntries.length, 1);
  for (let i = 0; i < maxFlavPay; i++) {
    const bg = i % 2 === 0 ? CREAM : WHITE;
    if (flavourEntries[i])  leftRow(flavourEntries[i][0],  flavourEntries[i][1]  + ' orders', bg);
    else if (!flavourEntries.length  && i === 0) leftRow('No data yet', '', CREAM);
    if (paymentEntries[i]) rightRow(paymentEntries[i][0], paymentEntries[i][1] + ' orders', bg);
    else if (!paymentEntries.length && i === 0) rightRow('No data yet', '', CREAM);
    row++;
  }

  spacer(20);

  // ══════════════════════════════════════
  //  CUSTOMER CRM (full width)
  // ══════════════════════════════════════
  banner('👥  Customer CRM', MOCHA, 13);

  // Reset to CRM-friendly column widths
  dash.setColumnWidth(1, 150);
  dash.setColumnWidth(2, 195);
  dash.setColumnWidth(3, 145);
  dash.setColumnWidth(4, 105);
  dash.setColumnWidth(5, 215);
  dash.setColumnWidth(6, 55);
  dash.setColumnWidth(7, 105);
  // col 8 for Notes — add it
  dash.setColumnWidth(8, 185);

  const crmHeaders = ['Name','Email','Phone / Instagram','Pickup Date','Flavours','Qty','Payment','Notes'];
  crmHeaders.forEach((h, i) => {
    dash.getRange(row, i + 1)
      .setValue(h)
      .setBackground(OLIVE)
      .setFontColor(WHITE)
      .setFontWeight('bold')
      .setFontSize(10)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  });
  dash.setRowHeight(row, 30);
  const crmHeaderRow = row;
  row++;

  if (orders.length === 0) {
    dash.getRange(row, 1, 1, 8).merge()
      .setValue("No orders yet — they're coming!")
      .setBackground(CREAM).setFontColor(BROWN)
      .setHorizontalAlignment('center').setFontStyle('italic');
    dash.setRowHeight(row, 28);
    row++;
  } else {
    orders.forEach((r_, i) => {
      const bg = i % 2 === 0 ? CREAM : WHITE;
      // Skip r_[0] (timestamp) — start from r_[1] = Name
      [r_[1], r_[2], r_[3], r_[4], r_[5], r_[6], r_[7], r_[8]].forEach((val, col) => {
        dash.getRange(row, col + 1)
          .setValue(val)
          .setBackground(bg)
          .setFontColor(BROWN)
          .setVerticalAlignment('middle')
          .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
      });
      dash.setRowHeight(row, 24);
      row++;
    });
  }

  dash.setFrozenRows(crmHeaderRow);

  // Last updated
  spacer(8);
  dash.getRange(row, 1, 1, 8).merge()
    .setValue('Last updated: ' + new Date().toLocaleString())
    .setFontStyle('italic').setFontColor('#AAAAAA')
    .setHorizontalAlignment('right').setBackground(WHITE);
}

function topEntry(obj) {
  let top = null, max = 0;
  Object.entries(obj).forEach(([k,v]) => { if (v > max) { max = v; top = k; } });
  return top;
}

// ─────────────────────────────────────────────────────────────────
//  TRIGGER SETUP — run once manually, then dashboard auto-rebuilds
//  on every change to the spreadsheet
// ─────────────────────────────────────────────────────────────────
function setupTrigger() {
  // Remove any existing onChange triggers to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'buildDashboard') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // Create a new onChange trigger
  ScriptApp.newTrigger('buildDashboard')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onChange()
    .create();
  Logger.log('Trigger created — dashboard will auto-rebuild on every change.');
}

// ─────────────────────────────────────────────────────────────────
//  SETUP — run once manually to format the Orders sheet
// ─────────────────────────────────────────────────────────────────
function setupSheet(sheet) {
  const headers = ['Timestamp','Name','Email','Phone / Instagram','Pickup Date','Flavours','Quantity','Payment','Notes'];
  sheet.appendRow(headers);

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange
    .setBackground('#4A5E2A').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  sheet.setRowHeight(1, 40);
  [160,140,200,160,120,240,150,110,260].forEach((w,i) => sheet.setColumnWidth(i+1, w));
  sheet.setFrozenRows(1);
  headerRange.setBorder(true,true,true,true,true,true,'#3A4A20', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.setTabColor('#4A5E2A');
  sheet.setName('Orders');
}
