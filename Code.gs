// ============================================================
//  ORDER MANAGEMENT SYSTEM — Code.gs
//  Paste this into your Google Apps Script editor
// ============================================================

// ─── CONFIGURATION ──────────────────────────────────────────
// Change these to match your business
var BUSINESS_NAME        = 'Burmelin';          // Your business name
var EMPLOYEE_PASS        = 'employee2024';         // Password for employee dashboard
var ORDERS_TAB_NAME      = 'Orders';               // Name for the new orders tab
var INVENTORY_SHEET_NAME = 'Size Chart';            // Your inventory tab name
var RETAIL_PASSWORDS = {
  'Bicasso 1':    'b1store',
  'Bicasso 2':    'b2store',
  'Bicasso Nana': 'bNana'
};
// ────────────────────────────────────────────────────────────

// ─── WEB APP ENTRY POINT ─────────────────────────────────────
// Now doubles as a JSON API for the Vercel-hosted frontend.
//   GET  ?action=getBusinessName | getInventory     → JSON
//   POST {action:'getOrders'|'submitOrder'|...}      → JSON
// Falls back to serving the old HTML pages if no action is given.
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';
  if (action) return handleApi(action, e.parameter, null);

  // Old Apps Script links are disabled — redirect everyone to the new Vercel site
  return HtmlService.createHtmlOutput(
    '<meta http-equiv="refresh" content="0; url=https://stockorder-form.vercel.app">'
    + '<style>body{font-family:Arial,sans-serif;display:flex;align-items:center;justify-content:center;min-height:100vh;margin:0;background:#f0f2f5}'
    + '.box{text-align:center;padding:40px;background:#fff;border-radius:16px;box-shadow:0 2px 20px rgba(0,0,0,.1)}'
    + 'h2{color:#1a1a2e;margin-bottom:10px}p{color:#666;margin-bottom:20px}'
    + 'a{background:#4f46e5;color:#fff;padding:12px 28px;border-radius:8px;text-decoration:none;font-weight:700}'
    + '</style>'
    + '<div class="box"><h2>This link has moved</h2>'
    + '<p>Please use the new link below.</p>'
    + '<a href="https://stockorder-form.vercel.app">Go to Burmelin Order Form</a></div>'
  ).setTitle('Moved').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// POST entry — body is JSON: { action: '...', ...payload }
function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) {}
  return handleApi(body.action, body, body);
}

// Wrap any value as a JSON response (readable cross-origin by the Vercel site)
function jsonOut(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}

// Central API router — maps an action name to the existing functions
function handleApi(action, params, body) {
  var out;
  switch (action) {
    case 'getBusinessName': out = BUSINESS_NAME; break;
    case 'getInventory':    out = getInventory(); break;
    case 'getOrders':       out = getOrders(params.password); break;
    case 'submitOrder':     out = submitOrder(body.order); break;
    case 'sendInvoice':     out = sendInvoice(body.orderId, body.password); break;
    case 'deleteOrder':     out = deleteOrder(body.orderId, body.password); break;
    case 'editOrderItems':  out = editOrderItems(body.orderId, body.newItems, body.password); break;
    case 'markDone':             out = markDone(body.orderId, body.password); break;
    case 'updateStockQty':       out = updateStockQty(body.productId, body.colorNum, body.size, body.newQty, body.password); break;
    case 'getAllRetailStock':     out = getAllRetailStock(); break;
    case 'getStoreStock':        out = getRetailStock(body && body.storeName); break;
    case 'sellRetailStock':      out = sellRetailStock(body.storeName, body.password, body.rowId, body.qty); break;
    case 'transferStock':        out = transferStock(body.fromStore, body.toStore, body.password, body.rowId, body.qty); break;
    case 'verifyRetailStore':    out = verifyRetailStore(body.storeName, body.password); break;
    case 'updateRetailStockQty': out = updateRetailStockQty(body.storeName, body.password, body.productId, body.colorNum, body.size, body.newQty); break;
    case 'addRetailStockItem':   out = addRetailStockItem(body.storeName, body.password, body.productId, body.category, body.colorNum, body.size, body.qty); break;
    case 'addRetailStockItems':    out = addRetailStockItems(body.storeName, body.password, body.items); break;
    case 'generateMasterSheet':  out = generateMasterSheet(body.password); break;
    default:                     out = { success: false, error: 'Unknown action: ' + action };
  }
  return jsonOut(out);
}

// ─── GET BUSINESS NAME (called by HTML pages) ─────────────────
function getBusinessName() {
  return BUSINESS_NAME;
}

// ─── GET INVENTORY (called by Customer form on load) ──────────
// Reads your sheet live so the form always shows current stock.
// Sheet columns: A=ID, B=ProductID, C=Category, D=Color, E=Size, F=Stock
function getInventory() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = null;

    // Find the inventory sheet by name (trim in case of trailing spaces)
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getName().trim() === INVENTORY_SHEET_NAME.trim()) {
        sheet = sheets[i]; break;
      }
    }
    // Fallback: use first non-Orders sheet
    if (!sheet) {
      for (var j = 0; j < sheets.length; j++) {
        if (sheets[j].getName() !== ORDERS_TAB_NAME) { sheet = sheets[j]; break; }
      }
    }
    if (!sheet) return { success: false, error: 'Inventory sheet not found.' };

    var data       = sheet.getDataRange().getValues();
    var inventory  = {};   // productId → { "Color X": { size: stock } }
    var categories = {};   // productId → category label

    // Find the header row (the row containing "Product ID" in column B)
    var dataStart = 1;
    for (var h = 0; h < data.length; h++) {
      if (String(data[h][1]).trim().toLowerCase() === 'product id') {
        dataStart = h + 1; // start reading from the row after headers
        break;
      }
    }

    for (var r = dataStart; r < data.length; r++) {
      var row       = data[r];
      var productId = String(row[1]).trim();
      var category  = String(row[2]).trim();
      var color     = String(row[3]).trim();
      var size      = String(row[4]).trim();
      var stock     = parseInt(row[5]);
      if (isNaN(stock)) stock = 0;

      if (!productId || !color || !size) continue;

      var colorKey = 'Color ' + color;

      if (!inventory[productId]) {
        inventory[productId] = {};
        categories[productId] = category;
      }
      if (!inventory[productId][colorKey]) inventory[productId][colorKey] = {};
      inventory[productId][colorKey][size] = stock;
    }

    // Build groups from categories (preserves insertion order)
    var groupMap = {};
    var groupOrder = [];
    Object.keys(categories).forEach(function(pid) {
      var cat = categories[pid];
      if (!groupMap[cat]) { groupMap[cat] = []; groupOrder.push(cat); }
      groupMap[cat].push(pid);
    });
    var groups = groupOrder.map(function(cat) {
      return { label: cat, products: groupMap[cat] };
    });

    return { success: true, inventory: inventory, groups: groups };

  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ─── SUBMIT ORDER (called from Customer form) ─────────────────
function submitOrder(orderData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ordersSheet = ss.getSheetByName(ORDERS_TAB_NAME);

    // Create Orders tab if it doesn't exist
    if (!ordersSheet) {
      ordersSheet = ss.insertSheet(ORDERS_TAB_NAME);
      var headers = [['Order ID','Date','Customer Name','Customer Email',
                      'Customer Phone','Notes','Items (JSON)','Total Qty',
                      'Status','Invoice Sent At','Completed At','Order Type','Shop Name']];
      ordersSheet.getRange(1, 1, 1, 13).setValues(headers);
      ordersSheet.getRange(1, 1, 1, 13).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
      ordersSheet.setFrozenRows(1);
      ordersSheet.setColumnWidth(7, 300); // Items column wider
    }

    // Generate order ID: ORD-YYMMDD-XXXX
    var now    = new Date();
    var datePart = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyMMdd');
    var rand   = Math.floor(1000 + Math.random() * 9000);
    var orderId = 'ORD-' + datePart + '-' + rand;
    var dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');

    var items      = orderData.items || [];
    var itemsJson  = JSON.stringify(items);
    var totalQty   = items.reduce(function(sum, i) { return sum + (i.qty || 0); }, 0);

    // Append order row
    ordersSheet.appendRow([
      orderId,
      dateStr,
      orderData.customerName  || '',
      orderData.customerEmail || '',
      orderData.customerPhone || '',
      orderData.notes         || '',
      itemsJson,
      totalQty,
      'Pending',
      '',  // Invoice sent at
      '',  // Completed at
      orderData.orderType || 'customer',
      orderData.shopName  || ''
    ]);

    // Color-code the new row
    var lastRow = ordersSheet.getLastRow();
    ordersSheet.getRange(lastRow, 9).setBackground('#fef9c3'); // Pending = yellow

    // Send auto-confirmation to customer
    try { sendConfirmationEmail(orderData, orderId, dateStr); } catch(e) {}

    // Notify the business owner
    try { notifyOwner(orderData, orderId, dateStr); } catch(e) {}

    return { success: true, orderId: orderId };

  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ─── GET ALL ORDERS (called from Employee dashboard) ──────────
function getOrders(password) {
  if (password !== EMPLOYEE_PASS) {
    return { success: false, error: 'Wrong password. Please try again.' };
  }
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(ORDERS_TAB_NAME);

    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, orders: [] };
    }

    var numRows = sheet.getLastRow() - 1;
    var numCols = sheet.getLastColumn();
    var data    = sheet.getRange(2, 1, numRows, numCols).getValues();

    var orders = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue; // skip empty rows
      orders.push({
        orderId      : row[0],
        date         : row[1],
        customerName : row[2],
        customerEmail: row[3],
        customerPhone: row[4],
        notes        : row[5],
        items        : JSON.parse(row[6] || '[]'),
        totalQty     : row[7],
        status       : row[8] || 'Pending',
        invoiceSent  : row[9],
        completed    : row[10],
        orderType    : row[11] || 'customer',
        shopName     : row[12] || ''
      });
    }

    // Sort: Pending first, then Invoiced, then Completed
    var order = { 'Pending': 0, 'Invoiced': 1, 'Completed': 2 };
    orders.sort(function(a, b) {
      return (order[a.status] || 0) - (order[b.status] || 0);
    });

    return { success: true, orders: orders };

  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ─── MARK ORDER AS INVOICED (called after employee prints invoice) ────
function sendInvoice(orderId, password) {
  if (password !== EMPLOYEE_PASS) {
    return { success: false, error: 'Wrong password.' };
  }
  try {
    var order = findOrder(orderId);
    if (!order) return { success: false, error: 'Order not found: ' + orderId };

    var sentAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    updateOrderRow(orderId, { status: 'Invoiced', invoiceSent: sentAt });

    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ─── DELETE ORDER (called from Employee dashboard) ────────────
function deleteOrder(orderId, password) {
  if (password !== EMPLOYEE_PASS) {
    return { success: false, error: 'Wrong password.' };
  }
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(ORDERS_TAB_NAME);
    if (!sheet || sheet.getLastRow() < 2) return { success: false, error: 'No orders found.' };

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]) === String(orderId)) {
        sheet.deleteRow(i + 2);
        return { success: true };
      }
    }
    return { success: false, error: 'Order not found.' };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ─── EDIT ORDER ITEMS (called from Employee dashboard) ───────
function editOrderItems(orderId, newItems, password) {
  if (password !== EMPLOYEE_PASS) {
    return { success: false, error: 'Wrong password.' };
  }
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(ORDERS_TAB_NAME);
    if (!sheet || sheet.getLastRow() < 2) return { success: false, error: 'Orders sheet not found.' };

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]) === String(orderId)) {
        var r        = i + 2;
        var totalQty = newItems.reduce(function(s, item) { return s + (item.qty || 0); }, 0);
        sheet.getRange(r, 7).setValue(JSON.stringify(newItems));
        sheet.getRange(r, 8).setValue(totalQty);
        return { success: true, totalQty: totalQty };
      }
    }
    return { success: false, error: 'Order not found.' };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ─── MARK ORDER DONE (called from Employee dashboard) ─────────
function markDone(orderId, password) {
  if (password !== EMPLOYEE_PASS) {
    return { success: false, error: 'Wrong password.' };
  }
  try {
    var order = findOrder(orderId);
    if (!order) return { success: false, error: 'Order not found: ' + orderId };

    // ── GUARD: if already completed, do nothing (prevents double deduction) ──
    if (order.status === 'Completed') {
      return { success: true, warnings: [], alreadyCompleted: true };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var warnings = [];

    if (order.orderType === 'store' && order.shopName) {
      // Store order: add ordered items to the shop's stock sheet
      for (var j = 0; j < order.items.length; j++) {
        var sit = order.items[j];
        var lp  = sit.lp  || sit.product || '';
        var lc  = sit.lc  || (sit.color || '').replace(/^Color\s*/i, '');
        var ls  = sit.ls  || sit.size    || '';
        var sres = addStockToStore(ss, order.shopName, lp, lc, ls, sit.qty || 0);
        if (!sres.success) {
          warnings.push(sit.product + ' | ' + sit.color + ' | ' + sit.size + ': ' + sres.error);
        }
      }
    } else {
      // Customer order: deduct stock from wholesale Size Chart
      for (var i = 0; i < order.items.length; i++) {
        var item   = order.items[i];
        var result = deductStock(ss, item);
        if (!result.success) {
          warnings.push(item.product + ' | ' + item.color + ' | ' + item.size + ': ' + result.error);
        }
      }
    }

    // Mark as Completed — also record the invoice timestamp if not already set
    var doneAt  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    var updates = { status: 'Completed', completed: doneAt };
    if (!order.invoiceSent) updates.invoiceSent = doneAt;
    updateOrderRow(orderId, updates);

    // Save final confirmed items to Sent Orders sheet
    try { saveSentOrder(ss, order, doneAt); } catch(e) {}

    return { success: true, warnings: warnings };

  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ─── UPDATE STOCK QTY (called from Employee stock view) ──────
function updateStockQty(productId, colorNum, size, newQty, password) {
  if (password !== EMPLOYEE_PASS) return { success: false, error: 'Wrong password.' };
  try {
    var ss     = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();

    for (var s = 0; s < sheets.length; s++) {
      var sheet = sheets[s];
      if (sheet.getName().trim() === ORDERS_TAB_NAME) continue;
      var data = sheet.getDataRange().getValues();

      var startRow = 1;
      for (var h = 0; h < data.length; h++) {
        if (String(data[h][1]).trim().toLowerCase() === 'product id') { startRow = h + 1; break; }
      }

      for (var r = startRow; r < data.length; r++) {
        if (String(data[r][1]).trim() === String(productId).trim() &&
            String(data[r][3]).trim() === String(colorNum).trim()  &&
            String(data[r][4]).trim() === String(size).trim()) {
          sheet.getRange(r + 1, 6).setValue(parseInt(newQty) || 0);
          return { success: true };
        }
      }
    }
    return { success: false, error: 'Item not found in sheet.' };
  } catch(err) {
    return { success: false, error: err.toString() };
  }
}

// ─── INTERNAL: SAVE SENT ORDER ────────────────────────────────
// Records the final employee-confirmed items to a "Sent Orders" sheet.
// This reflects what was actually sent from the warehouse, not the original request.
function saveSentOrder(ss, order, doneAt) {
  var SENT_SHEET = 'Sent Orders';
  var sheet = ss.getSheetByName(SENT_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(SENT_SHEET);
    var headers = [['Order ID','Date Sent','Customer Name','Customer Phone','Notes','Product','Color','Size','Qty']];
    sheet.getRange(1, 1, 1, 9).setValues(headers);
    sheet.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, 9, 120);
  }

  var items = order.items || [];
  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    sheet.appendRow([
      order.orderId,
      doneAt,
      order.customerName  || '',
      order.customerPhone || '',
      order.notes         || '',
      item.product,
      item.color,
      item.size,
      item.qty
    ]);
    // Alternate row shading
    var lastRow = sheet.getLastRow();
    if (lastRow % 2 === 0) sheet.getRange(lastRow, 1, 1, 9).setBackground('#f8fafc');
  }
}

// ─── INTERNAL: FIND ORDER ROW ─────────────────────────────────
function findOrder(orderId) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ORDERS_TAB_NAME);
  if (!sheet || sheet.getLastRow() < 2) return null;

  var numCols = sheet.getLastColumn();
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, numCols).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(orderId)) {
      return {
        rowIndex     : i + 2,
        orderId      : data[i][0],
        date         : data[i][1],
        customerName : data[i][2],
        customerEmail: data[i][3],
        customerPhone: data[i][4],
        notes        : data[i][5],
        items        : JSON.parse(data[i][6] || '[]'),
        totalQty     : data[i][7],
        status       : data[i][8],
        invoiceSent  : data[i][9],
        orderType    : data[i][11] || 'customer',
        shopName     : data[i][12] || ''
      };
    }
  }
  return null;
}

// ─── INTERNAL: UPDATE ORDER ROW ───────────────────────────────
function updateOrderRow(orderId, updates) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ORDERS_TAB_NAME);
  if (!sheet) return;

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(orderId)) {
      var r = i + 2; // actual sheet row

      if (updates.status) {
        sheet.getRange(r, 9).setValue(updates.status);
        // Color-code status
        var bg = updates.status === 'Completed' ? '#dcfce7'
               : updates.status === 'Invoiced'  ? '#dbeafe'
               : '#fef9c3';
        sheet.getRange(r, 9).setBackground(bg);
      }
      if (updates.invoiceSent !== undefined) sheet.getRange(r, 10).setValue(updates.invoiceSent);
      if (updates.completed   !== undefined) sheet.getRange(r, 11).setValue(updates.completed);
      return;
    }
  }
}

// ─── INTERNAL: ADD STOCK TO RETAIL STORE (used when marking store order Done) ─
function addStockToStore(ss, shopName, productId, colorNum, size, qty) {
  try {
    var sheetName = shopName + ' Stock';
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, error: 'Sheet not found for store: ' + shopName };

    var data = sheet.getDataRange().getValues();
    var startRow = 1;
    for (var h = 0; h < data.length; h++) {
      if (String(data[h][1]).trim().toLowerCase() === 'product id') { startRow = h + 1; break; }
    }

    var pid = String(productId).trim();
    var col = String(colorNum).trim();
    var sz  = String(size).trim();

    for (var r = startRow; r < data.length; r++) {
      if (String(data[r][1]).trim() === pid &&
          String(data[r][3]).trim() === col &&
          String(data[r][4]).trim() === sz) {
        var cur = parseInt(data[r][5]) || 0;
        sheet.getRange(r + 1, 6).setValue(cur + (parseInt(qty) || 0));
        return { success: true };
      }
    }
    // Row doesn't exist — create it
    var compositeId = pid + col + sz;
    sheet.appendRow([compositeId, pid, '', col, sz, parseInt(qty) || 0]);
    return { success: true };
  } catch(err) {
    return { success: false, error: err.toString() };
  }
}

// ─── INTERNAL: DEDUCT STOCK ───────────────────────────────────
// Sheet columns: A=ID, B=ProductID, C=Category, D=Color, E=Size, F=Stock
// item.lp = ProductID, item.lc = Color number, item.ls = Size
function deductStock(ss, item) {
  try {
    var sheets = ss.getSheets();

    for (var s = 0; s < sheets.length; s++) {
      var sheet = sheets[s];
      if (sheet.getName().trim() === ORDERS_TAB_NAME) continue;
      if (sheet.getLastRow() < 2) continue;

      var data = sheet.getDataRange().getValues();

      // Find where data actually starts (skip title + header rows)
      var startRow = 1;
      for (var h = 0; h < data.length; h++) {
        if (String(data[h][1]).trim().toLowerCase() === 'product id') {
          startRow = h + 1; break;
        }
      }

      for (var r = startRow; r < data.length; r++) {
        var row = data[r];
        var productId = String(row[1]).trim();
        var color     = String(row[3]).trim();
        var size      = String(row[4]).trim();

        // Use lp/lc/ls (lookup fields) when available; fall back to product/color/size
        var matchProd  = String(item.lp  || item.product || '').trim();
        var matchColor = String(item.lc  || (item.color  || '').replace(/^Color\s*/i, '')).trim();
        var matchSize  = String(item.ls  || item.size    || '').trim();

        if (productId === matchProd &&
            color     === matchColor &&
            size      === matchSize) {
          var currentStock = parseInt(row[5]) || 0;
          var newStock     = Math.max(0, currentStock - (item.qty || 0));
          sheet.getRange(r + 1, 6).setValue(newStock);
          return { success: true };
        }
      }
    }

    return { success: false, error: 'Item not found in sheet' };

  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ─── INTERNAL: CONFIRMATION EMAIL TO CUSTOMER ─────────────────
function sendConfirmationEmail(orderData, orderId, dateStr) {
  if (!orderData.customerEmail) return;

  var itemLines = (orderData.items || []).map(function(i) {
    return '  • ' + i.product + '  |  ' + i.color + '  |  Size ' + i.size + '  |  Qty: ' + i.qty;
  }).join('\n');

  var body = 'Hi ' + orderData.customerName + ',\n\n'
    + 'Thank you! We received your order.\n\n'
    + 'Order ID : ' + orderId + '\n'
    + 'Date     : ' + dateStr + '\n\n'
    + 'Items:\n' + itemLines + '\n\n'
    + (orderData.notes ? 'Notes: ' + orderData.notes + '\n\n' : '')
    + 'We will prepare your order and send the invoice shortly.\n\n'
    + BUSINESS_NAME;

  GmailApp.sendEmail(
    orderData.customerEmail,
    'Order Received — ' + orderId + ' | ' + BUSINESS_NAME,
    body,
    { name: BUSINESS_NAME }
  );
}

// ─── INTERNAL: NOTIFY OWNER OF NEW ORDER ──────────────────────
function notifyOwner(orderData, orderId, dateStr) {
  var ownerEmail = Session.getActiveUser().getEmail();
  if (!ownerEmail) return;

  var itemLines = (orderData.items || []).map(function(i) {
    return '  • ' + i.product + '  |  ' + i.color + '  |  Size ' + i.size + '  |  Qty: ' + i.qty;
  }).join('\n');

  GmailApp.sendEmail(
    ownerEmail,
    '🛒 New Order: ' + orderId + ' from ' + orderData.customerName,
    'New order received!\n\n'
    + 'Order ID : ' + orderId + '\n'
    + 'Date     : ' + dateStr + '\n'
    + 'Customer : ' + orderData.customerName + '\n'
    + 'Email    : ' + orderData.customerEmail + '\n'
    + 'Phone    : ' + (orderData.customerPhone || 'N/A') + '\n\n'
    + 'Items:\n' + itemLines + '\n\n'
    + (orderData.notes ? 'Notes: ' + orderData.notes + '\n\n' : '')
    + 'Open your Employee Dashboard to manage this order.',
    { name: BUSINESS_NAME + ' – Order System' }
  );
}

// ─── TRANSFER STOCK BETWEEN STORES ───────────────────────────
function transferStock(fromStore, toStore, password, rowId, qty) {
  if (!fromStore || !toStore) return { success: false, error: 'Select both stores.' };
  if (fromStore === toStore) return { success: false, error: 'Source and destination must differ.' };
  var sellQ = parseInt(qty) || 0;
  if (sellQ <= 0) return { success: false, error: 'Quantity must be at least 1.' };
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    function sheetFor(name) {
      return ss.getSheetByName(name === 'Palladium' ? INVENTORY_SHEET_NAME : name + ' Stock');
    }
    var fromSheet = sheetFor(fromStore);
    var toSheet   = sheetFor(toStore);
    if (!fromSheet) return { success: false, error: 'Sheet not found for: ' + fromStore };
    if (!toSheet)   return { success: false, error: 'Sheet not found for: ' + toStore };
    function findRow(sheet) {
      var data = sheet.getDataRange().getValues();
      for (var r = 1; r < data.length; r++) {
        if (String(data[r][0]).trim().toUpperCase() === String(rowId).trim().toUpperCase())
          return { row: r + 1, qty: parseInt(data[r][5]) || 0 };
      }
      return null;
    }
    var from = findRow(fromSheet);
    if (!from) return { success: false, error: 'ID not found in ' + fromStore + ': ' + rowId };
    if (from.qty < sellQ) return { success: false, error: 'Not enough stock in ' + fromStore + '. Available: ' + from.qty };
    var to = findRow(toSheet);
    if (!to) return { success: false, error: 'ID not found in ' + toStore + ': ' + rowId };
    fromSheet.getRange(from.row, 6).setValue(from.qty - sellQ);
    toSheet.getRange(to.row, 6).setValue(to.qty + sellQ);
    return { success: true, fromNewQty: from.qty - sellQ, toNewQty: to.qty + sellQ };
  } catch(e) { return { success: false, error: e.toString() }; }
}

// ─── SELL FROM RETAIL STORE ───────────────────────────────────
function sellRetailStock(storeName, password, rowId, qty) {
  var validPass = (RETAIL_PASSWORDS[storeName] && password === RETAIL_PASSWORDS[storeName])
               || password === EMPLOYEE_PASS;
  if (!validPass) return { success: false, error: 'Wrong password.' };
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(storeName + ' Stock');
    if (!sheet) return { success: false, error: 'Store sheet not found.' };
    var data  = sheet.getDataRange().getValues();
    var sellQ = parseInt(qty) || 0;
    if (sellQ <= 0) return { success: false, error: 'Quantity must be at least 1.' };
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][0]).trim().toUpperCase() === String(rowId).trim().toUpperCase()) {
        var cur    = parseInt(data[r][5]) || 0;
        if (cur < sellQ) return { success: false, error: 'Not enough stock. Available: ' + cur };
        var newQty = cur - sellQ;
        sheet.getRange(r + 1, 6).setValue(newQty);
        return { success: true, newQty: newQty };
      }
    }
    return { success: false, error: 'ID not found: ' + rowId };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// ─── VERIFY RETAIL STORE LOGIN ────────────────────────────────
function verifyRetailStore(storeName, password) {
  if (!RETAIL_PASSWORDS[storeName]) return { success: false, error: 'Unknown store.' };
  if (password !== RETAIL_PASSWORDS[storeName]) return { success: false, error: 'Wrong password.' };
  return { success: true };
}

// ─── GET ALL RETAIL STOCK (public read, no password needed) ───
function getAllRetailStock() {
  return getRetailStock(null);
}

// ─── GET RETAIL STOCK (one store or all stores) ───────────────
function getRetailStock(storeName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var storeNames = storeName ? [storeName] : ['Bicasso 1', 'Bicasso 2', 'Bicasso Nana', 'Palladium'];
    var storesData = {};

    storeNames.forEach(function(name) {
      var sheetName = name === 'Palladium' ? INVENTORY_SHEET_NAME : name + ' Stock';
      var sheet = ss.getSheetByName(sheetName);
      // Auto-create sheet from wholesale product list if missing (not for Palladium)
      if (!sheet && name !== 'Palladium') sheet = createRetailSheet(ss, sheetName);
      var inventory = {};

      if (sheet && sheet.getLastRow() > 1) {
        var data = sheet.getDataRange().getValues();
        var startRow = 1;
        for (var h = 0; h < data.length; h++) {
          if (String(data[h][1]).trim().toLowerCase() === 'product id') { startRow = h + 1; break; }
        }
        for (var r = startRow; r < data.length; r++) {
          var row      = data[r];
          var pid      = String(row[1]).trim();
          var color    = String(row[3]).trim();
          var size     = String(row[4]).trim();
          var stock    = parseInt(row[5]) || 0;
          if (!pid || !color || !size) continue;
          var colorKey = 'Color ' + color;
          if (!inventory[pid]) inventory[pid] = {};
          if (!inventory[pid][colorKey]) inventory[pid][colorKey] = {};
          inventory[pid][colorKey][size] = stock;
        }
      }
      storesData[name] = inventory;
    });

    return { success: true, stores: storesData };
  } catch(err) {
    return { success: false, error: err.toString() };
  }
}

// ─── CREATE RETAIL STORE SHEET (seeded from wholesale, qty=0) ─
function createRetailSheet(ss, sheetName) {
  try {
    var wholesale = null;
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getName().trim() === INVENTORY_SHEET_NAME.trim()) { wholesale = sheets[i]; break; }
    }
    var newSheet = ss.insertSheet(sheetName);
    if (wholesale) {
      var src  = wholesale.getDataRange().getValues();
      var dest = src.map(function(row) { return row.slice(0, 6); });
      // Zero out stock column (F) for all data rows
      var startRow = 1;
      for (var h = 0; h < dest.length; h++) {
        if (String(dest[h][1]).trim().toLowerCase() === 'product id') { startRow = h + 1; break; }
      }
      for (var r = startRow; r < dest.length; r++) { dest[r][5] = 0; }
      newSheet.getRange(1, 1, dest.length, 6).setValues(dest);
      newSheet.getRange(startRow, 1, 1, 6).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
      newSheet.setFrozenRows(startRow);
    } else {
      var headers = [['ID','Product ID','Category','Color','Size','Stock']];
      newSheet.getRange(1, 1, 1, 6).setValues(headers);
      newSheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
      newSheet.setFrozenRows(1);
    }
    return newSheet;
  } catch(e) {
    return null;
  }
}

// ─── ADD MULTIPLE ROWS TO RETAIL STOCK SHEET (batch) ────────
// items = [{productId, category, colorNum, size, qty}, ...]
function addRetailStockItems(storeName, password, items) {
  var validPass = (RETAIL_PASSWORDS[storeName] && password === RETAIL_PASSWORDS[storeName])
               || password === EMPLOYEE_PASS;
  if (!validPass) return { success: false, error: 'Wrong password.' };
  try {
    var ss        = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = storeName + ' Stock';
    var sheet     = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, error: 'Sheet for "' + storeName + '" not found.' };

    var data     = sheet.getDataRange().getValues();
    var startRow = 1;
    for (var h = 0; h < data.length; h++) {
      if (String(data[h][1]).trim().toLowerCase() === 'product id') { startRow = h + 1; break; }
    }

    var results = [];
    for (var n = 0; n < items.length; n++) {
      var it = items[n];
      var pid = String(it.productId).trim();
      var col = String(it.colorNum).trim();
      var sz  = String(it.size).trim();
      var qty = parseInt(it.qty) || 0;
      var found = false;
      // Re-read data each iteration so previous appends are visible
      var d2 = sheet.getDataRange().getValues();
      for (var r = startRow; r < d2.length; r++) {
        if (String(d2[r][1]).trim() === pid &&
            String(d2[r][3]).trim() === col &&
            String(d2[r][4]).trim() === sz) {
          sheet.getRange(r + 1, 6).setValue(qty);
          found = true; break;
        }
      }
      if (!found) {
        var compositeId = pid + col + sz;
        sheet.appendRow([compositeId, pid, it.category || '', col, sz, qty]);
      }
      results.push({ size: sz, ok: true });
    }
    return { success: true, results: results };
  } catch(err) {
    return { success: false, error: err.toString() };
  }
}

// ─── ADD NEW ROW TO RETAIL STOCK SHEET ───────────────────────
// Palladium/employee only — appends a new product row to a store's sheet
function addRetailStockItem(storeName, password, productId, category, colorNum, size, qty) {
  var validPass = (RETAIL_PASSWORDS[storeName] && password === RETAIL_PASSWORDS[storeName])
               || password === EMPLOYEE_PASS;
  if (!validPass) return { success: false, error: 'Wrong password.' };
  try {
    var ss        = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = storeName + ' Stock';
    var sheet     = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, error: 'Sheet for "' + storeName + '" not found.' };

    var data     = sheet.getDataRange().getValues();
    // Check it doesn't already exist
    var startRow = 1;
    for (var h = 0; h < data.length; h++) {
      if (String(data[h][1]).trim().toLowerCase() === 'product id') { startRow = h + 1; break; }
    }
    for (var r = startRow; r < data.length; r++) {
      if (String(data[r][1]).trim() === String(productId).trim() &&
          String(data[r][3]).trim() === String(colorNum).trim()  &&
          String(data[r][4]).trim() === String(size).trim()) {
        // Row exists — update qty instead of failing
        sheet.getRange(r + 1, 6).setValue(parseInt(qty) || 0);
        return { success: true };
      }
    }
    // Composite ID: ProductID + Color + Size (e.g. T9027 + 91 + 48 = T9027914 8)
    var compositeId = String(productId).trim() + String(colorNum).trim() + String(size).trim();
    sheet.appendRow([compositeId, productId, category || '', colorNum, size, parseInt(qty) || 0]);
    return { success: true };
  } catch(err) {
    return { success: false, error: err.toString() };
  }
}

// ─── UPDATE RETAIL STOCK QTY ──────────────────────────────────
// Accepts store-specific passcode OR master employee password
function updateRetailStockQty(storeName, password, productId, colorNum, size, newQty) {
  var validPass = (RETAIL_PASSWORDS[storeName] && password === RETAIL_PASSWORDS[storeName])
               || password === EMPLOYEE_PASS;
  if (!validPass) return { success: false, error: 'Wrong password.' };
  try {
    var ss        = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = storeName + ' Stock';
    var sheet     = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, error: 'Sheet for "' + storeName + '" not found. Please contact admin.' };

    var data     = sheet.getDataRange().getValues();
    var startRow = 1;
    for (var h = 0; h < data.length; h++) {
      if (String(data[h][1]).trim().toLowerCase() === 'product id') { startRow = h + 1; break; }
    }
    for (var r = startRow; r < data.length; r++) {
      if (String(data[r][1]).trim() === String(productId).trim() &&
          String(data[r][3]).trim() === String(colorNum).trim()  &&
          String(data[r][4]).trim() === String(size).trim()) {
        sheet.getRange(r + 1, 6).setValue(parseInt(newQty) || 0);
        return { success: true };
      }
    }
    return { success: false, error: 'Item not found in ' + storeName + ' stock.' };
  } catch(err) {
    return { success: false, error: err.toString() };
  }
}

// ─── INTERNAL: BUILD INVOICE HTML ─────────────────────────────
function buildInvoiceHtml(order) {
  var rows = (order.items || []).map(function(item) {
    return '<tr>'
      + '<td style="padding:12px 14px;border-bottom:1px solid #f0f0f0">' + item.product + '</td>'
      + '<td style="padding:12px 14px;border-bottom:1px solid #f0f0f0">' + item.color   + '</td>'
      + '<td style="padding:12px 14px;border-bottom:1px solid #f0f0f0">' + item.size    + '</td>'
      + '<td style="padding:12px 14px;border-bottom:1px solid #f0f0f0;text-align:center;font-weight:bold">' + item.qty + '</td>'
      + '</tr>';
  }).join('');

  return '<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body style="margin:0;padding:30px;background:#f5f5f5;font-family:Arial,sans-serif">'
  + '<div style="max-width:600px;margin:0 auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 2px 20px rgba(0,0,0,0.1)">'
  + '<div style="background:#1a1a2e;color:#fff;padding:35px 30px;text-align:center">'
  +   '<h1 style="margin:0;font-size:1.6rem;font-weight:700">' + BUSINESS_NAME + '</h1>'
  +   '<p style="margin:8px 0 0;opacity:0.75;font-size:0.95rem">Order Invoice</p>'
  + '</div>'
  + '<div style="padding:30px">'
  +   '<table style="width:100%;margin-bottom:25px;font-size:0.9rem;color:#555">'
  +     '<tr><td style="padding:5px 0;color:#999;width:130px">Order ID</td><td style="font-weight:bold;color:#1a1a2e">' + order.orderId + '</td></tr>'
  +     '<tr><td style="padding:5px 0;color:#999">Date</td><td>' + order.date + '</td></tr>'
  +     '<tr><td style="padding:5px 0;color:#999">Customer</td><td><strong>' + order.customerName + '</strong></td></tr>'
  +     (order.customerPhone ? '<tr><td style="padding:5px 0;color:#999">Phone</td><td>' + order.customerPhone + '</td></tr>' : '')
  +   '</table>'
  +   '<h3 style="font-size:0.85rem;text-transform:uppercase;letter-spacing:0.05em;color:#999;margin-bottom:8px">Order Items</h3>'
  +   '<table style="width:100%;border-collapse:collapse;font-size:0.9rem">'
  +     '<thead><tr style="background:#f8fafc">'
  +       '<th style="padding:10px 14px;text-align:left;font-size:0.78rem;text-transform:uppercase;color:#aaa">Product</th>'
  +       '<th style="padding:10px 14px;text-align:left;font-size:0.78rem;text-transform:uppercase;color:#aaa">Color</th>'
  +       '<th style="padding:10px 14px;text-align:left;font-size:0.78rem;text-transform:uppercase;color:#aaa">Size</th>'
  +       '<th style="padding:10px 14px;text-align:center;font-size:0.78rem;text-transform:uppercase;color:#aaa">Qty</th>'
  +     '</tr></thead>'
  +     '<tbody>' + rows + '</tbody>'
  +   '</table>'
  +   '<div style="background:#f8fafc;border-radius:8px;padding:16px 14px;margin-top:16px;display:flex;justify-content:space-between;align-items:center">'
  +     '<span style="color:#555;font-size:0.9rem">Total Items</span>'
  +     '<span style="font-size:1.1rem;font-weight:700;color:#1a1a2e">' + order.totalQty + '</span>'
  +   '</div>'
  +   (order.notes ? '<div style="margin-top:16px;padding:14px;background:#fffbeb;border-radius:8px;border-left:3px solid #f59e0b;font-size:0.9rem"><strong>Notes:</strong> ' + order.notes + '</div>' : '')
  + '</div>'
  + '<div style="background:#f8fafc;padding:22px 30px;text-align:center;color:#aaa;font-size:0.85rem">'
  +   '<p style="margin:0">Thank you for your order! Questions? Just reply to this email.</p>'
  + '</div>'
  + '</div></body></html>';
}

// ─── GENERATE MASTER SHEET ───────────────────────────────────────
// Creates/refreshes a "Master" tab with full stock overview:
// Product ID | Color | Size | Palladium | Bicasso 1 | Bicasso 2 | Bicasso Nana | Total | Status
function generateMasterSheet(password) {
  if (password !== EMPLOYEE_PASS) return { success: false, error: 'Wrong password.' };
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var MASTER_NAME = 'Master';
    var master = ss.getSheetByName(MASTER_NAME);
    if (!master) {
      master = ss.insertSheet(MASTER_NAME);
    } else {
      master.clearContents();
      master.clearFormats();
    }

    var RETAIL_STORES = ['Bicasso 1', 'Bicasso 2', 'Bicasso Nana'];
    var headers = ['Product ID', 'Color', 'Size', 'Palladium', 'Bicasso 1', 'Bicasso 2', 'Bicasso Nana', 'Total Network', 'Status'];
    master.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
    master.setFrozenRows(1);

    // Load wholesale (Palladium) stock from Size Chart
    var wholesaleSheet = null;
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getName().trim() === INVENTORY_SHEET_NAME.trim()) { wholesaleSheet = sheets[i]; break; }
    }
    if (!wholesaleSheet) return { success: false, error: 'Wholesale sheet not found.' };

    var wData = wholesaleSheet.getDataRange().getValues();
    var wStart = 1;
    for (var h = 0; h < wData.length; h++) {
      if (String(wData[h][1]).trim().toLowerCase() === 'product id') { wStart = h + 1; break; }
    }

    // Load each retail store stock into a lookup map
    var retailMap = {};
    RETAIL_STORES.forEach(function(storeName) {
      retailMap[storeName] = {};
      var sheet = ss.getSheetByName(storeName + ' Stock');
      if (!sheet) return;
      var data = sheet.getDataRange().getValues();
      var start = 1;
      for (var h = 0; h < data.length; h++) {
        if (String(data[h][1]).trim().toLowerCase() === 'product id') { start = h + 1; break; }
      }
      for (var r = start; r < data.length; r++) {
        var pid = String(data[r][1]).trim();
        var cn  = String(data[r][3]).trim();
        var sz  = String(data[r][4]).trim();
        var qty = parseInt(data[r][5]) || 0;
        if (!pid || !cn || !sz) continue;
        var key = pid + '||' + cn + '||' + sz;
        retailMap[storeName][key] = qty;
      }
    });

    // Build data rows from wholesale reference
    var rows = [];
    var bgColors = [];
    for (var r = wStart; r < wData.length; r++) {
      var row = wData[r];
      var pid    = String(row[1]).trim();
      var cn     = String(row[3]).trim();
      var sz     = String(row[4]).trim();
      var palQty = parseInt(row[5]) || 0;
      if (!pid || !cn || !sz) continue;

      var key = pid + '||' + cn + '||' + sz;
      var b1  = retailMap['Bicasso 1'][key]    || 0;
      var b2  = retailMap['Bicasso 2'][key]    || 0;
      var bn  = retailMap['Bicasso Nana'][key] || 0;
      var total = palQty + b1 + b2 + bn;
      var status = palQty === 0 ? 'Order Now' : palQty <= 3 ? 'Low' : 'OK';

      rows.push([pid, cn, sz, palQty, b1, b2, bn, total, status]);
      bgColors.push(status === 'Order Now' ? '#fee2e2' : status === 'Low' ? '#fef9c3' : '#f0fdf4');
    }

    if (rows.length) {
      master.getRange(2, 1, rows.length, headers.length).setValues(rows);
      // Color-code the Status and Palladium columns
      for (var i = 0; i < rows.length; i++) {
        var bg = bgColors[i];
        master.getRange(i + 2, 9).setBackground(bg).setFontWeight('bold');
        master.getRange(i + 2, 4).setBackground(bg);
      }
    }

    master.autoResizeColumns(1, headers.length);
    var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    return { success: true, rows: rows.length, updated: dateStr };

  } catch(err) {
    return { success: false, error: err.toString() };
  }
}
