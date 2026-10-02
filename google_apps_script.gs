// ============================================================
// GALLOPS MENSTRUAL CUP - ORDER MANAGEMENT (Google Apps Script)
// Version 2 - Handles both GET params and POST JSON
// ============================================================

var SHEET_NAME = 'Orders';
var NOTIFY_EMAIL = 'earthenterprise100@gmail.com';

function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    var headers = ['Order ID','Date','Name','Mobile','Email','Address','Landmark','City','State','PIN Code','Size','Qty','Total (Rs)','Payment','Notes','How Heard','Status'];
    sheet.appendRow(headers);
    var hRange = sheet.getRange(1, 1, 1, headers.length);
    hRange.setBackground('#f06292');
    hRange.setFontColor('#ffffff');
    hRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 120);
    sheet.setColumnWidth(3, 150);
    sheet.setColumnWidth(5, 180);
    sheet.setColumnWidth(6, 200);
  }
  return sheet;
}

function saveOrderToSheet(data) {
  var sheet = getOrCreateSheet();
  sheet.appendRow([
    data.id, data.date, data.name, data.mobile, data.email,
    data.address, data.landmark, data.city, data.state, data.pincode,
    data.size, data.qty, data.total, data.payment, data.notes,
    data.howHeard, 'New'
  ]);
  sendEmailNotification(data);
}

function sendEmailNotification(data) {
  try {
    var subject = 'New Gallops Order ' + data.id + ' - ' + data.name + ' - Rs.' + data.total;
    var body = '*** NEW ORDER RECEIVED ***\n\n'
      + 'Order ID: ' + data.id + '\n'
      + 'Date: ' + data.date + '\n\n'
      + '--- CUSTOMER DETAILS ---\n'
      + 'Name: ' + data.name + '\n'
      + 'Mobile: ' + data.mobile + '\n'
      + 'Email: ' + data.email + '\n\n'
      + '--- DELIVERY ADDRESS ---\n'
      + 'Address: ' + data.address + '\n'
      + 'Landmark: ' + data.landmark + '\n'
      + 'City: ' + data.city + '\n'
      + 'State: ' + data.state + '\n'
      + 'PIN: ' + data.pincode + '\n\n'
      + '--- ORDER DETAILS ---\n'
      + 'Product: Gallops ' + data.size + ' Menstrual Cup\n'
      + 'Quantity: ' + data.qty + '\n'
      + 'Total Amount: Rs.' + data.total + '\n'
      + 'Payment Method: ' + data.payment + '\n'
      + (data.notes ? 'Customer Notes: ' + data.notes + '\n' : '')
      + (data.howHeard ? 'Heard via: ' + data.howHeard + '\n' : '')
      + '\n--- ACTION REQUIRED ---\n'
      + 'Please confirm this order by calling/WhatsApp on: ' + data.mobile + '\n'
      + 'View all orders in your Google Sheet.';
    MailApp.sendEmail(NOTIFY_EMAIL, subject, body);
  } catch(err) {
    Logger.log('Email error: ' + err);
  }
}

// Handle GET requests (save order via URL params OR fetch order list)
function doGet(e) {
  var output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    var params = e.parameter;

    // Save order action
    if (params.action === 'save') {
      var data = {
        id:        params.id || '',
        date:      params.date || new Date().toLocaleString(),
        name:      params.name || '',
        mobile:    params.mobile || '',
        email:     params.email || '',
        address:   params.address || '',
        landmark:  params.landmark || '',
        city:      params.city || '',
        state:     params.state || '',
        pincode:   params.pincode || '',
        size:      params.size || '',
        qty:       params.qty || 1,
        total:     params.total || 0,
        payment:   params.payment || '',
        notes:     params.notes || '',
        howHeard:  params.howHeard || '',
        status:    'New'
      };
      saveOrderToSheet(data);
      output.setContent(JSON.stringify({success: true, id: data.id}));
      return output;
    }

    // Default: return all orders
    var sheet = getOrCreateSheet();
    var rows = sheet.getDataRange().getValues();
    var orders = [];
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      if (!r[0]) continue;
      orders.push({
        id: r[0], date: r[1], name: r[2], mobile: r[3], email: r[4],
        address: r[5], landmark: r[6], city: r[7], state: r[8], pincode: r[9],
        size: r[10], qty: r[11], total: r[12], payment: r[13], notes: r[14],
        howHeard: r[15], status: r[16]
      });
    }
    orders.reverse();
    output.setContent(JSON.stringify({success: true, orders: orders}));
    return output;

  } catch(err) {
    output.setContent(JSON.stringify({success: false, error: err.toString()}));
    return output;
  }
}

// Handle POST requests (JSON body)
function doPost(e) {
  var output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  try {
    var data = JSON.parse(e.postData.contents);
    saveOrderToSheet(data);
    output.setContent(JSON.stringify({success: true, id: data.id}));
  } catch(err) {
    output.setContent(JSON.stringify({success: false, error: err.toString()}));
  }
  return output;
}
