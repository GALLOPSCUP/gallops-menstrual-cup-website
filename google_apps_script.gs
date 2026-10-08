// ============================================================
// GALLOPS MENSTRUAL CUP - ORDER MANAGEMENT
// Version 6 - Supports New Orders, Status Update, Courier & Delete
// Mail order confirmation alert removed as per user request
// ============================================================

var SHEET_NAME = 'Orders';

function cors(output) {
  return output;
}

function getOrCreateSheet() {
  var props = PropertiesService.getScriptProperties();
  var sheetId = props.getProperty('SHEET_ID');
  var ss;
  
  if (sheetId) {
    try {
      ss = SpreadsheetApp.openById(sheetId);
    } catch(e) {
      sheetId = null;
    }
  }
  
  if (!sheetId) {
    ss = SpreadsheetApp.create("Gallops Orders Database");
    props.setProperty('SHEET_ID', ss.getId());
  }
  
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  if (sheet.getName() !== SHEET_NAME) {
    sheet.setName(SHEET_NAME);
  }
  
  if (sheet.getLastRow() === 0) {
    var h = ['Order ID','Date','Name','Mobile','Email','Address',
             'Landmark','City','State','PIN','Size','Qty','Total (Rs)',
             'Payment','Notes','How Heard','Status','Courier','Tracking ID'];
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
  var orderDate = d.date;
  if (!orderDate) {
    orderDate = Utilities.formatDate(new Date(), 'Asia/Kolkata', "yyyy-MM-dd'T'HH:mm:ss.SSS'Z'");
  }
  sheet.appendRow([
    d.id, orderDate, d.name, d.mobile, d.email,
    d.address, d.landmark, d.city, d.state, d.pincode,
    d.size, d.qty, d.total, d.payment,
    d.notes || '', d.howHeard || '', d.status || 'New',
    d.courier || '', d.trackingNo || ''
  ]);
  // Mail order confirmation alert deleted as per user request
}

function updateOrderCourier(id, courier, trackingNo) {
  var sheet = getOrCreateSheet();
  var rows = sheet.getDataRange().getValues();
  if (rows[0].length < 18 || !rows[0][17]) sheet.getRange(1, 18).setValue('Courier');
  if (rows[0].length < 19 || !rows[0][18]) sheet.getRange(1, 19).setValue('Tracking ID');

  var targetId = String(id).trim();
  var updated = false;
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === targetId) {
      if (courier !== undefined && courier !== null) sheet.getRange(i + 1, 18).setValue(courier);
      if (trackingNo !== undefined && trackingNo !== null) sheet.getRange(i + 1, 19).setValue(trackingNo);
      updated = true;
    }
  }
  return updated;
}

function updateOrderStatus(id, newStatus) {
  var sheet = getOrCreateSheet();
  var rows = sheet.getDataRange().getValues();
  var targetId = String(id).trim();
  var updated = false;
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === targetId) {
      sheet.getRange(i + 1, 17).setValue(newStatus);
      updated = true;
    }
  }
  return updated;
}

function deleteOrderById(id) {
  var sheet = getOrCreateSheet();
  var rows = sheet.getDataRange().getValues();
  var targetId = String(id).trim();
  var deleted = false;
  // Loop backwards so row indexes don't shift
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]).trim() === targetId) {
      sheet.deleteRow(i + 1);
      deleted = true;
    }
  }
  return deleted;
}

// ---------- GET - returns order list or handles status update / delete / courier ----------
function doGet(e) {
  var out = ContentService.createTextOutput();
  out.setMimeType(ContentService.MimeType.JSON);
  try {
    var params = (e && e.parameter) ? e.parameter : {};
    var action = params.action;
    var sheet  = getOrCreateSheet();

    if (action === 'updateCourier' && params.id) {
      var ok = updateOrderCourier(params.id, params.courier || '', params.trackingNo || '');
      out.setContent(JSON.stringify({ success: ok, id: params.id, courier: params.courier, trackingNo: params.trackingNo }));
      return out;
    }

    if (action === 'updateStatus' && params.id) {
      var ok = updateOrderStatus(params.id, params.status || 'Confirmed');
      out.setContent(JSON.stringify({ success: ok, id: params.id, status: params.status }));
      return out;
    }

    if (action === 'deleteOrder' && params.id) {
      var ok = deleteOrderById(params.id);
      out.setContent(JSON.stringify({ success: ok, id: params.id }));
      return out;
    }

    var rows   = sheet.getDataRange().getValues();
    var orders = [];
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      if (!r[0]) continue;
      var rawDate = r[1];
      var formattedDate = rawDate;
      if (rawDate instanceof Date) {
        formattedDate = rawDate.toISOString();
      }
      orders.push({
        id:r[0], date:formattedDate, name:r[2], mobile:r[3], email:r[4],
        address:r[5], landmark:r[6], city:r[7], state:r[8], pincode:r[9],
        size:r[10], qty:r[11], total:r[12], payment:r[13],
        notes:r[14], howHeard:r[15], status:r[16] || 'New', courier:r[17] || '', trackingNo:r[18] || ''
      });
    }
    out.setContent(JSON.stringify({ success:true, orders:orders.reverse() }));
  } catch(e) {
    out.setContent(JSON.stringify({ success:false, error:e.toString() }));
  }
  return out;
}

// ---------- POST - receives order, status update, delete, or courier ----------
function doPost(e) {
  var out = ContentService.createTextOutput();
  out.setMimeType(ContentService.MimeType.JSON);
  try {
    var raw  = e.postData.contents;
    var data = JSON.parse(raw);

    if (data.action === 'updateStatus' && data.id) {
      var ok = updateOrderStatus(data.id, data.status || 'Confirmed');
      out.setContent(JSON.stringify({ success: ok, id: data.id, status: data.status }));
      return out;
    }

    if (data.action === 'updateCourier' && data.id) {
      var ok = updateOrderCourier(data.id, data.courier || '', data.trackingNo || '');
      out.setContent(JSON.stringify({ success: ok, id: data.id, courier: data.courier, trackingNo: data.trackingNo }));
      return out;
    }

    if (data.action === 'deleteOrder' && data.id) {
      var ok = deleteOrderById(data.id);
      out.setContent(JSON.stringify({ success: ok, id: data.id }));
      return out;
    }

    saveOrder(data);
    out.setContent(JSON.stringify({ success:true, id:data.id }));
  } catch(e) {
    out.setContent(JSON.stringify({ success:false, error:e.toString() }));
  }
  return out;
}