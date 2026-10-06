const SPREADSHEET_ID = '1NUnsAEktAnoxA5FT4uGw_N9MvQYYTZAbAIzgQmS--FI'; 
// ផ្អែកលើរូបភាព ផ្ទាំងដែលកំពុងបើកមានឈ្មោះ "បញ្ជីម្ចាស់កឋិនទាន (1)"
const SHEET_NAME = 'បញ្ជីម្ចាស់កឋិនទាន (1)';

function doGet(e) {
  // If requested as JSON API
  if (e && e.parameter && e.parameter.api === 'true') {
    const data = getData();
    return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Otherwise, serve standard Web App HTML
  // Note: Vercel uses index.html, but Apps Script still uses 'Index' (the name of the file in the GAS editor)
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('បញ្ជីរាយនាមម្ចាស់អង្គកឋិនទាន')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Support POST requests from Vercel
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const method = payload.method;
    const args = payload.args || [];
    
    let result = null;
    
    if (method === 'getData') result = getData();
    else if (method === 'addData') result = addData(args[0]);
    else if (method === 'updateData') result = updateData(args[0], args[1]);
    else if (method === 'deleteData') result = deleteData(args[0]);
    else if (method === 'deleteMultipleData') result = deleteMultipleData(args[0]);
    else if (method === 'reorderData') result = reorderData(args[0], args[1], args[2]);
    else throw new Error('Method not found: ' + method);

    return ContentService.createTextOutput(JSON.stringify({ status: 'success', data: result }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getData() {
  try {
    const ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("មិនអាចស្វែងរកឯកសារ Spreadsheet បានទេ។ សូមពិនិត្យ SPREADSHEET_ID។");
    
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error("មិនអាចស្វែងរកផ្ទាំងឈ្មោះ '" + SHEET_NAME + "' បានទេ។");

    const data = sheet.getDataRange().getValues();
    if (data.length < 4) return { rows: [], updated: new Date().toLocaleString() };

    const rows = [];
    for (let i = 4; i < data.length; i++) {
      const row = data[i];
      const id = row[0]; // No.
      const name = row[1]; // Donor name
      
      if (id !== "" && !isNaN(id) && name !== "") {
        rows.push({
          id: id,
          name: name ? String(name).trim() : "",
          village: row[2] ? String(row[2]).trim() : "",
          construction: parseFloat(row[3]) || 0,
          con_curr: row[4] ? String(row[4]).trim() : "",
          requisites: parseFloat(row[5]) || 0,
          req_curr: row[6] ? String(row[6]).trim() : "",
          meal: parseFloat(row[7]) || 0,
          meal_curr: row[8] ? String(row[8]).trim() : "",
          type: row[9] ? String(row[9]).trim() : "", 
          notes: row[10] ? String(row[10]).trim() : ""
        });
      }
    }
    
    const timeZone = ss.getSpreadsheetTimeZone();
    const updated = Utilities.formatDate(new Date(), timeZone, "dd/MM/yyyy HH:mm:ss");
    
    return { rows: rows, updated: updated };
  } catch (error) {
    throw new Error(error.message);
  }
}

function addData(record) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    
    let lastValidIndex = -1;
    let maxId = 0;
    
    for (let i = 4; i < data.length; i++) {
      const id = data[i][0];
      const name = data[i][1];
      if (id !== "" && !isNaN(id) && name !== "") {
        lastValidIndex = i;
        if (Number(id) > maxId) maxId = Number(id);
      }
    }
    
    const newId = maxId + 1;
    const insertRowPos = (lastValidIndex !== -1) ? lastValidIndex + 2 : 5; 
    
    sheet.insertRowAfter(insertRowPos - 1);
    
    const newRow = [
      newId,
      record.name || "",
      record.village || "",
      record.construction || "",
      record.con_curr || "",
      record.requisites || "",
      record.req_curr || "",
      record.meal || "",
      record.meal_curr || "",
      record.type || "",
      record.notes || ""
    ];
    
    sheet.getRange(insertRowPos, 1, 1, 11).setValues([newRow]);
    
    // Copy format from the row above if it exists and is a valid data row
    if (insertRowPos > 5) {
      const rangeToCopy = sheet.getRange(insertRowPos - 1, 1, 1, 11);
      rangeToCopy.copyTo(sheet.getRange(insertRowPos, 1, 1, 11), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      // Also copy data validations (dropdowns)
      rangeToCopy.copyTo(sheet.getRange(insertRowPos, 1, 1, 11), SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    }
    
    return getData();
  } catch (error) {
    throw new Error(error.message);
  }
}

function updateData(id, record) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    
    let targetRow = -1;
    for (let i = 4; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(id).trim()) {
        targetRow = i + 1;
        break;
      }
    }
    
    if (targetRow === -1) throw new Error("មិនរកឃើញទិន្នន័យដែលត្រូវកែប្រែទេ!");
    
    const updatedRow = [
      id,
      record.name || "",
      record.village || "",
      record.construction || "",
      record.con_curr || "",
      record.requisites || "",
      record.req_curr || "",
      record.meal || "",
      record.meal_curr || "",
      record.type || "",
      record.notes || ""
    ];
    
    sheet.getRange(targetRow, 1, 1, 11).setValues([updatedRow]);
    return getData();
  } catch (error) {
    throw new Error(error.message);
  }
}

function deleteData(id) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    
    let targetRow = -1;
    let availableIds = [];
    for (let i = 4; i < data.length; i++) {
      availableIds.push(String(data[i][0]).trim());
      if (String(data[i][0]).trim() === String(id).trim()) {
        targetRow = i + 1;
        break;
      }
    }
    
    if (targetRow === -1) {
      throw new Error("មិនរកឃើញទិន្នន័យដែលត្រូវលុបទេ! (IDដែលបញ្ជូនមក: '" + id + "', IDមានក្នុងតារាង: [" + availableIds.join(", ") + "])");
    }
    
    sheet.deleteRow(targetRow);
    return getData();
  } catch (error) {
    throw new Error(error.message);
  }
}

function deleteMultipleData(ids) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    
    const rowsToDelete = [];
    const idList = ids.map(id => String(id).trim());
    for (let i = 4; i < data.length; i++) {
      if (idList.includes(String(data[i][0]).trim())) {
        rowsToDelete.push(i + 1);
      }
    }
    
    rowsToDelete.sort((a, b) => b - a);
    
    for (let i = 0; i < rowsToDelete.length; i++) {
      sheet.deleteRow(rowsToDelete[i]);
    }
    
    return getData();
  } catch (error) {
    throw new Error(error.message);
  }
}

// មុខងារសម្រាប់បញ្ចូលទិន្នន័យសាកល្បង ៥០ នាក់
function addTestData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  
  const data = sheet.getDataRange().getValues();
  let maxId = 0;
  for (let i = 4; i < data.length; i++) {
    const currentId = parseInt(data[i][0], 10);
    if (!isNaN(currentId) && currentId > maxId) {
      maxId = currentId;
    }
  }
  
  const rows = [];
  const villages = ['តាំងគោក', 'សូយោង', 'ភ្នំពេញ', 'សៀមរាប'];
  const types = ['ត្រៃ', 'លៀង', 'ស្បង់', ''];
  
  for (let i = 1; i <= 50; i++) {
    const newId = maxId + i;
    const v = villages[Math.floor(Math.random() * villages.length)];
    const t = types[Math.floor(Math.random() * types.length)];
    
    rows.push([
      newId,
      "ឈ្មោះសាកល្បង ទី" + newId,
      v,
      Math.floor(Math.random() * 100) * 1000,
      "រៀល",
      Math.floor(Math.random() * 50) * 1000,
      "រៀល",
      Math.floor(Math.random() * 50),
      "ដុល្លារ",
      t,
      "ទិន្នន័យតេស្ត"
    ]);
  }
  
  const startRow = data.length < 4 ? 5 : data.length + 1;
  sheet.getRange(startRow, 1, rows.length, 11).setValues(rows);
}

function reorderData(draggedId, targetId, position) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    
    let draggedRowIdx = -1;
    let targetRowIdx = -1;
    
    for (let i = 4; i < data.length; i++) {
      const id = String(data[i][0]).trim();
      if (id === String(draggedId)) draggedRowIdx = i + 1;
      if (id === String(targetId)) targetRowIdx = i + 1;
    }
    
    if (draggedRowIdx !== -1 && targetRowIdx !== -1 && draggedRowIdx !== targetRowIdx) {
      const rangeToMove = sheet.getRange(draggedRowIdx, 1, 1, sheet.getLastColumn());
      
      let destIndex = targetRowIdx;
      if (position === 'after') {
        destIndex += 1;
      }
      
      // If the destination index resolves to the exact same physical position,
      // moveRows throws "The destination index cannot be within the span being moved."
      // For a single row (size 1), destIndex === draggedRowIdx (moving to its own start)
      // or destIndex === draggedRowIdx + 1 (moving to immediately after itself) are no-ops.
      if (destIndex !== draggedRowIdx && destIndex !== draggedRowIdx + 1) {
        // Move the row to the destination
        sheet.moveRows(rangeToMove, destIndex);
        
        // Wait for moveRows to complete then renumber all IDs
        SpreadsheetApp.flush();
        
        const lastRow = sheet.getLastRow();
        if (lastRow >= 5) {
          const newIds = [];
          for (let i = 5; i <= lastRow; i++) {
            newIds.push([i - 4]); // Resets IDs to 1, 2, 3...
          }
          sheet.getRange(5, 1, newIds.length, 1).setValues(newIds);
        }
      }
    }
    
    return getData();
  } catch (error) {
    throw new Error(error.message);
  }
}
