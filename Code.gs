// ============================================================
//  ORDER MANAGEMENT SYSTEM — Code.gs
//  Paste this into your Google Apps Script editor
// ============================================================

// ─── CONFIGURATION ──────────────────────────────────────────
// Change these to match your business
var BUSINESS_NAME        = 'Burmelin';          // Your business name
var EMPLOYEE_PASS        = 'employee2024';         // Password for employee dashboard
var ORDERS_TAB_NAME      = 'Orders';               // Name for the new orders tab
var INVENTORY_SHEET_NAME = 'Updated Size Chart';   // Your inventory tab name
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
    case 'markDone':        out = markDone(body.orderId, body.password); break;
    default:                out = { success: false, error: 'Unknown action: ' + action };
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
                      'Status','Invoice Sent At','Completed At']];
      ordersSheet.getRange(1, 1, 1, 11).setValues(headers);
      ordersSheet.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
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
      ''   // Completed at
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
    var data    = sheet.getRange(2, 1, numRows, 11).getValues();

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
        completed    : row[10]
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

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var warnings = [];

    // Deduct stock for each item
    for (var i = 0; i < order.items.length; i++) {
      var item   = order.items[i];
      var result = deductStock(ss, item);
      if (!result.success) {
        warnings.push(item.product + ' | ' + item.color + ' | ' + item.size + ': ' + result.error);
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

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getValues();
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
        status       : data[i][8]
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

        if (productId === String(item.lp).trim() &&
            color     === String(item.lc).trim() &&
            size      === String(item.ls).trim()) {
          var currentStock = parseInt(row[5]) || 0;
          var newStock     = Math.max(0, currentStock - (item.qty || 0));
          sheet.getRange(r + 1, 6).setValue(newStock); // col F = column 6
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
