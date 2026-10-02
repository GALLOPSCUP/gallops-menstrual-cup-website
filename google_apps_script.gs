// ============================================================
// GALLOPS MENSTRUAL CUP - ORDER MANAGEMENT (Google Apps Script)
// INSTRUCTIONS:
// 1. Go to script.google.com  -> New Project
// 2. Paste this entire code
// 3. Click Deploy -> New Deployment -> Web App
// 4. Set "Execute as" = Me, "Who has access" = Anyone
// 5. Click Deploy -> Copy the Web App URL
// 6. Paste that URL in order.html where it says GOOGLE_SCRIPT_URL
// ============================================================

var SHEET_NAME = 'Orders';
var NOTIFY_EMAIL = 'earthenterprise100@gmail.com'; // Your email

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    
    // Create sheet + headers if not exists
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow([
        'Order ID','Date','Name','Mobile','Email',
        'Address','Landmark','City','State','PIN Code',
        'Size','Qty','Total (Rs)','Payment','Notes',
        'How Heard','Status'
      ]);
      sheet.getRange(1,1,1,17).setBackground('#f06292').setFontColor('#fff').setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    
    // Append order row
    sheet.appendRow([
      data.id, data.date, data.name, data.mobile, data.email,
      data.address, data.landmark, data.city, data.state, data.pincode,
      data.size, data.qty, data.total, data.payment, data.notes,
      data.howHeard, 'New'
    ]);
    
    // Send email notification
    var subject = 'New Gallops Order ' + data.id + ' from ' + data.name;
    var body = 'NEW ORDER RECEIVED!\n\n'
      + 'Order ID: ' + data.id + '\n'
      + 'Date: ' + data.date + '\n\n'
      + '--- CUSTOMER ---\n'
      + 'Name: ' + data.name + '\n'
      + 'Mobile: ' + data.mobile + '\n'
      + 'Email: ' + data.email + '\n\n'
      + '--- DELIVERY ADDRESS ---\n'
      + data.address + '\n'
      + 'Landmark: ' + data.landmark + '\n'
      + data.city + ', ' + data.state + ' - ' + data.pincode + '\n\n'
      + '--- ORDER ---\n'
      + 'Product: Gallops ' + data.size + ' Menstrual Cup\n'
      + 'Quantity: ' + data.qty + '\n'
      + 'Total: Rs.' + data.total + '\n'
      + 'Payment: ' + data.payment + '\n'
      + (data.notes ? 'Notes: ' + data.notes + '\n' : '')
      + '\nView all orders: https://docs.google.com/spreadsheets/d/'
      + SpreadsheetApp.getActiveSpreadsheet().getId();
    
    MailApp.sendEmail(NOTIFY_EMAIL, subject, body);
    
    return ContentService.createTextOutput(JSON.stringify({success:true, id:data.id}))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({success:false, error:err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return ContentService.createTextOutput(JSON.stringify({orders:[]})).setMimeType(ContentService.MimeType.JSON);
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var orders = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      orders.push({
        id:       row[0],  date:    row[1],  name:     row[2],
        mobile:   row[3],  email:   row[4],  address:  row[5],
        landmark: row[6],  city:    row[7],  state:    row[8],
        pincode:  row[9],  size:    row[10], qty:      row[11],
        total:    row[12], payment: row[13], notes:    row[14],
        howHeard: row[15], status:  row[16]
      });
    }
    
    return ContentService.createTextOutput(JSON.stringify({orders:orders.reverse()}))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({orders:[], error:err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function updateStatus(orderId, newStatus) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === orderId) {
      sheet.getRange(i+1, 17).setValue(newStatus);
      return true;
    }
  }
  return false;
}
