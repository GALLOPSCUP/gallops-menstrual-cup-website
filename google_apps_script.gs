// ============================================================
// GALLOPS MENSTRUAL CUP - ORDER MANAGEMENT
// Version 3 - Works with text/plain POST (no CORS preflight)
// ============================================================

var SHEET_NAME   = 'Orders';
var NOTIFY_EMAIL = 'earthenterprise100@gmail.com';

// ---------- helpers ----------
function cors(output) {
  return output; // Apps Script handles CORS automatically for deployed Web Apps
}

function getOrCreateSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    var h = ['Order ID','Date','Name','Mobile','Email','Address',
             'Landmark','City','State','PIN','Size','Qty','Total (Rs)',
             'Payment','Notes','How Heard','Status'];
    sheet.appendRow(h);
    var r = sheet.getRange(1, 1, 1, h.length);
    r.setBackground('#f06292');
    r.setFontColor('#fff');
    r.setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function saveOrder(d) {
  var sheet = getOrCreateSheet();
  sheet.appendRow([
    d.id, d.date, d.name, d.mobile, d.email,
    d.address, d.landmark, d.city, d.state, d.pincode,
    d.size, d.qty, d.total, d.payment,
    d.notes || '', d.howHeard || '', 'New'
  ]);
  sendMail(d);
}

function sendMail(d) {
  try {
    var sub  = 'New Gallops Order ' + d.id + ' | ' + d.name + ' | Rs.' + d.total;
    var body = '*** NEW ORDER ***\n\n'
      + 'Order ID : ' + d.id   + '\n'
      + 'Date     : ' + d.date + '\n\n'
      + 'CUSTOMER\n'
      + 'Name     : ' + d.name   + '\n'
      + 'Mobile   : ' + d.mobile + '\n'
      + 'Email    : ' + d.email  + '\n\n'
      + 'DELIVERY ADDRESS\n'
      + 'Address  : ' + d.address  + '\n'
      + 'Landmark : ' + d.landmark + '\n'
      + 'City     : ' + d.city     + '\n'
      + 'State    : ' + d.state    + '\n'
      + 'PIN      : ' + d.pincode  + '\n\n'
      + 'ORDER\n'
      + 'Product  : Gallops ' + d.size + ' Menstrual Cup\n'
      + 'Qty      : ' + d.qty     + '\n'
      + 'Total    : Rs.' + d.total + '\n'
      + 'Payment  : ' + d.payment  + '\n'
      + (d.notes ? 'Notes    : ' + d.notes + '\n' : '')
      + '\nPlease confirm with customer on: ' + d.mobile;
    MailApp.sendEmail(NOTIFY_EMAIL, sub, body);
  } catch(e) {
    Logger.log('Mail error: ' + e);
  }
}

// ---------- GET - returns order list ----------
function doGet(e) {
  var out = ContentService.createTextOutput();
  out.setMimeType(ContentService.MimeType.JSON);
  try {
    var sheet  = getOrCreateSheet();
    var rows   = sheet.getDataRange().getValues();
    var orders = [];
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      if (!r[0]) continue;
      orders.push({
        id:r[0], date:r[1], name:r[2], mobile:r[3], email:r[4],
        address:r[5], landmark:r[6], city:r[7], state:r[8], pincode:r[9],
        size:r[10], qty:r[11], total:r[12], payment:r[13],
        notes:r[14], howHeard:r[15], status:r[16]
      });
    }
    out.setContent(JSON.stringify({ success:true, orders:orders.reverse() }));
  } catch(e) {
    out.setContent(JSON.stringify({ success:false, error:e.toString() }));
  }
  return out;
}

// ---------- POST - receives order (text/plain body containing JSON) ----------
function doPost(e) {
  var out = ContentService.createTextOutput();
  out.setMimeType(ContentService.MimeType.JSON);
  try {
    var raw  = e.postData.contents;
    var data = JSON.parse(raw);
    saveOrder(data);
    out.setContent(JSON.stringify({ success:true, id:data.id }));
  } catch(e) {
    out.setContent(JSON.stringify({ success:false, error:e.toString() }));
  }
  return out;
}
