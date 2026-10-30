/**
 * ระบบตัดเบิกของใช้สำนักงานและแม่บ้าน (Office & Housekeeper Inventory Requisition System)
 * Google Apps Script Web App Backend
 */

const SHEET_CONFIGS = {
  "lc": {
    tabName: "อปก.คข.LC ",
    title: "อุปกรณ์-เครื่องเขียน LC (หลักชัย)",
    type: "stationery",
    nameCol: 2,   // Col B
    stockCol: 3,  // Col C (จำนวน MAX/คงเหลือ)
    minCol: 4,    // Col D
    unitCol: 5,   // Col E
    startRow: 2
  },
  "nll": {
    tabName: "อปก.คข.NLL",
    title: "อุปกรณ์-เครื่องเขียน NLL (หนองละลอก)",
    type: "stationery",
    nameCol: 2,   // Col B
    stockCol: 3,  // Col C
    minCol: 4,    // Col D
    unitCol: 5,   // Col E
    startRow: 2
  },
  "wc": {
    tabName: "อปก.คข.WC",
    title: "อุปกรณ์-เครื่องเขียน WC (วังจันทร์)",
    type: "stationery",
    nameCol: 2,   // Col B
    stockCol: 3,  // Col C
    minCol: 4,    // Col D
    unitCol: 5,   // Col E
    startRow: 2
  },
  "maid": {
    tabName: "อปก.แม่บ้าน ต.ค.",
    title: "อุปกรณ์แม่บ้าน",
    type: "maid",
    codeCol: 3,   // Col C (ลำดับ)
    nameCol: 4,   // Col D (รายการวัสดุ)
    stockCol: 5,  // Col E (จำนวน MAX/คงเหลือ)
    unitCol: 6,   // Col F (หน่วย)
    noteCol: 11,  // Col K (หมายเหตุ)
    startRow: 3
  }
};

const LOG_SHEET_REQUISITION = "ประวัติการเบิก";
const LOG_SHEET_RESTOCK = "ประวัติการเติมสต็อก";

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "getItems";
    
    if (action === "getItems") {
      const items = getAllItems();
      return jsonResponse({ success: true, data: items, timestamp: new Date().toISOString() });
    }
    
    if (action === "getLogs") {
      const logs = getRequisitionLogs();
      const restocks = getRestockLogs();
      return jsonResponse({ success: true, logs: logs, restocks: restocks });
    }
    
    if (action === "ping") {
      return jsonResponse({ success: true, message: "Office Stock API is active", time: new Date() });
    }
    
    return jsonResponse({ success: false, error: "Unknown action" });
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

function doPost(e) {
  try {
    let payload = {};
    if (e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents);
      } catch (ex) {
        payload = e.parameter || {};
      }
    } else {
      payload = e.parameter || {};
    }
    
    const action = payload.action;
    
    if (action === "requisition") {
      const result = handleRequisition(payload);
      return jsonResponse(result);
    }
    
    if (action === "restock") {
      const result = handleRestock(payload);
      return jsonResponse(result);
    }
    
    if (action === "addItem") {
      const result = handleAddItem(payload);
      return jsonResponse(result);
    }

    if (action === "updateItem") {
      const result = handleUpdateItem(payload);
      return jsonResponse(result);
    }
    
    return jsonResponse({ success: false, error: "Invalid action: " + action });
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

function getAllItems() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allItems = [];
  
  for (const catKey in SHEET_CONFIGS) {
    const cfg = SHEET_CONFIGS[catKey];
    const sheet = ss.getSheetByName(cfg.tabName);
    if (!sheet) continue;
    
    const data = sheet.getDataRange().getValues();
    
    for (let r = cfg.startRow - 1; r < data.length; r++) {
      const row = data[r];
      let name = "";
      let code = "";
      let stock = 0;
      let min = 0;
      let unit = "อัน";
      let note = "";
      
      if (cfg.type === "maid") {
        name = row[cfg.nameCol - 1] ? String(row[cfg.nameCol - 1]).trim() : "";
        code = row[cfg.codeCol - 1] ? String(row[cfg.codeCol - 1]).trim() : String(r);
        if (!name || name === "รายการวัสดุ" || name.indexOf("อุปกรณ์") !== -1) continue;
        
        const rawStock = row[cfg.stockCol - 1];
        stock = isNaN(Number(rawStock)) ? 0 : Number(rawStock);
        unit = row[cfg.unitCol - 1] ? String(row[cfg.unitCol - 1]).trim() : "อัน";
        note = row[cfg.noteCol - 1] ? String(row[cfg.noteCol - 1]).trim() : "";
      } else {
        const rawName = row[cfg.nameCol - 1] ? String(row[cfg.nameCol - 1]).trim() : "";
        if (!rawName || rawName.indexOf("เช็คของ") !== -1 || rawName === "วัสดุ อปก.") continue;
        
        name = rawName;
        const rawStock = row[cfg.stockCol - 1];
        stock = isNaN(Number(rawStock)) ? 0 : Number(rawStock);
        
        const rawMin = row[cfg.minCol - 1];
        min = (rawMin && !isNaN(Number(rawMin))) ? Number(rawMin) : 0;
        
        unit = row[cfg.unitCol - 1] ? String(row[cfg.unitCol - 1]).trim() : "อัน";
      }
      
      const cleanName = name.replace(/^\d+[\.\s]+/, "").trim();
      
      allItems.push({
        id: catKey + "_" + (r + 1),
        rowNumber: r + 1,
        category: catKey,
        categoryName: cfg.title,
        sheetTab: cfg.tabName,
        code: code || String(r + 1),
        name: cleanName,
        fullName: name,
        stock: stock,
        min: min,
        unit: unit,
        note: note
      });
    }
  }
  
  return allItems;
}

function handleRequisition(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requester = (payload.requester || "").trim();
  const items = payload.items || [];
  
  if (!requester) {
    return { success: false, error: "กรุณาระบุชื่อผู้เบิก" };
  }
  if (!items || items.length === 0) {
    return { success: false, error: "ไม่มีรายการที่เลือกเบิก" };
  }
  
  const now = new Date();
  const timestampStr = payload.timestamp || Utilities.formatDate(now, "GMT+7", "yyyy-MM-dd HH:mm:ss");
  const dateStr = payload.date || Utilities.formatDate(now, "GMT+7", "yyyy-MM-dd");
  const timeStr = payload.time || Utilities.formatDate(now, "GMT+7", "HH:mm");
  
  const logSheet = getOrCreateLogSheet(ss);
  const updatedItems = [];
  
  for (let i = 0; i < items.length; i++) {
    const itemReq = items[i];
    const catKey = itemReq.category;
    const cfg = SHEET_CONFIGS[catKey];
    if (!cfg) continue;
    
    const sheet = ss.getSheetByName(cfg.tabName);
    if (!sheet) continue;
    
    const reqQty = Number(itemReq.qty);
    if (isNaN(reqQty) || reqQty <= 0) continue;
    
    let targetRow = itemReq.rowNumber;
    const data = sheet.getDataRange().getValues();
    let found = false;
    
    if (targetRow && targetRow <= data.length) {
      const checkName = String(data[targetRow - 1][cfg.nameCol - 1] || "");
      if (checkName.indexOf(itemReq.name) !== -1 || itemReq.name.indexOf(checkName.replace(/^\d+[\.\s]+/, "").trim()) !== -1) {
        found = true;
      }
    }
    
    if (!found) {
      for (let r = cfg.startRow - 1; r < data.length; r++) {
        const curName = String(data[r][cfg.nameCol - 1] || "");
        const cleanCur = curName.replace(/^\d+[\.\s]+/, "").trim();
        if (cleanCur === itemReq.name || curName === itemReq.fullName) {
          targetRow = r + 1;
          found = true;
          break;
        }
      }
    }
    
    if (!found || !targetRow) continue;
    
    const curStockVal = sheet.getRange(targetRow, cfg.stockCol).getValue();
    const curStock = isNaN(Number(curStockVal)) ? 0 : Number(curStockVal);
    const newStock = Math.max(0, curStock - reqQty);
    
    sheet.getRange(targetRow, cfg.stockCol).setValue(newStock);
    
    // ใส่ ' นำหน้าเพื่อให้ Google Sheets เก็บเป็น Plain Text ไม่เพี้ยนเป็น Date object
    logSheet.appendRow([
      "'" + timestampStr,
      "'" + dateStr,
      "'" + timeStr,
      requester,
      cfg.title,
      itemReq.name,
      reqQty,
      itemReq.unit || "อัน",
      curStock,
      newStock,
      itemReq.note || ""
    ]);
    
    updatedItems.push({
      id: itemReq.id,
      category: catKey,
      name: itemReq.name,
      reqQty: reqQty,
      prevStock: curStock,
      newStock: newStock,
      unit: itemReq.unit
    });
  }
  
  return {
    success: true,
    message: "บันทึกการตัดเบิกเรียบร้อยแล้ว",
    timestamp: timestampStr,
    requester: requester,
    updatedItems: updatedItems
  };
}

function handleRestock(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const adminName = (payload.adminName || payload.requester || "ผู้ดูแลระบบ").trim();
  const items = payload.items || [];
  
  if (!items || items.length === 0) {
    return { success: false, error: "ไม่มีรายการที่เลือกเติมของ" };
  }
  
  const now = new Date();
  const timestampStr = payload.timestamp || Utilities.formatDate(now, "GMT+7", "yyyy-MM-dd HH:mm:ss");
  const dateStr = payload.date || Utilities.formatDate(now, "GMT+7", "yyyy-MM-dd");
  
  const restockSheet = getOrCreateRestockSheet(ss);
  const updatedItems = [];
  
  for (let i = 0; i < items.length; i++) {
    const itemRestock = items[i];
    const catKey = itemRestock.category;
    const cfg = SHEET_CONFIGS[catKey];
    if (!cfg) continue;
    
    const sheet = ss.getSheetByName(cfg.tabName);
    if (!sheet) continue;
    
    const addQty = Number(itemRestock.qty);
    if (isNaN(addQty) || addQty <= 0) continue;
    
    let targetRow = itemRestock.rowNumber;
    const data = sheet.getDataRange().getValues();
    let found = false;
    
    if (targetRow && targetRow <= data.length) {
      found = true;
    } else {
      for (let r = cfg.startRow - 1; r < data.length; r++) {
        const curName = String(data[r][cfg.nameCol - 1] || "");
        const cleanCur = curName.replace(/^\d+[\.\s]+/, "").trim();
        if (cleanCur === itemRestock.name || curName === itemRestock.fullName) {
          targetRow = r + 1;
          found = true;
          break;
        }
      }
    }
    
    if (!found || !targetRow) continue;
    
    const curStockVal = sheet.getRange(targetRow, cfg.stockCol).getValue();
    const curStock = isNaN(Number(curStockVal)) ? 0 : Number(curStockVal);
    const newStock = curStock + addQty;
    
    sheet.getRange(targetRow, cfg.stockCol).setValue(newStock);
    
    restockSheet.appendRow([
      "'" + timestampStr,
      "'" + dateStr,
      adminName,
      cfg.title,
      itemRestock.name,
      addQty,
      itemRestock.unit || "อัน",
      curStock,
      newStock,
      itemRestock.note || "เติมสต็อกโดยผู้ดูแล"
    ]);
    
    updatedItems.push({
      id: itemRestock.id,
      category: catKey,
      name: itemRestock.name,
      addQty: addQty,
      prevStock: curStock,
      newStock: newStock,
      unit: itemRestock.unit
    });
  }
  
  return {
    success: true,
    message: "บันทึกการเติมสต็อกเรียบร้อยแล้ว",
    timestamp: timestampStr,
    adminName: adminName,
    updatedItems: updatedItems
  };
}

/**
 * แก้ไขยอดสต็อกคงเหลือโดยตรง (Direct Stock Update by Admin)
 */
function handleUpdateItem(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const adminName = (payload.adminName || "ผู้ดูแลระบบ").trim();
  const catKey = payload.category;
  const cfg = SHEET_CONFIGS[catKey];
  if (!cfg) return { success: false, error: "ไม่พบหมวดหมู่นี้" };

  const sheet = ss.getSheetByName(cfg.tabName);
  if (!sheet) return { success: false, error: "ไม่พบชีต: " + cfg.tabName };

  const newStock = Number(payload.stock);
  if (isNaN(newStock) || newStock < 0) return { success: false, error: "จำนวนสต็อกไม่ถูกต้อง" };

  const data = sheet.getDataRange().getValues();
  let targetRow = payload.rowNumber;
  let found = false;

  if (targetRow && targetRow <= data.length) {
    found = true;
  } else {
    for (let r = cfg.startRow - 1; r < data.length; r++) {
      const curName = String(data[r][cfg.nameCol - 1] || "");
      const cleanCur = curName.replace(/^\d+[\.\s]+/, "").trim();
      if (cleanCur === payload.name || curName === payload.fullName) {
        targetRow = r + 1;
        found = true;
        break;
      }
    }
  }

  if (!found || !targetRow) return { success: false, error: "ไม่พบแถวของรายการใน Google Sheet" };

  const curStockVal = sheet.getRange(targetRow, cfg.stockCol).getValue();
  const prevStock = isNaN(Number(curStockVal)) ? 0 : Number(curStockVal);

  // อัปเดตยอดสต็อกใหม่ลงชีต
  sheet.getRange(targetRow, cfg.stockCol).setValue(newStock);

  // อัปเดตหน่วยนับ ถ้ามีการส่งมา
  if (payload.unit && cfg.unitCol) {
    sheet.getRange(targetRow, cfg.unitCol).setValue(payload.unit);
  }

  // บันทึกลงชีตประวัติการเติมสต็อก/ปรับปรุงยอด
  const restockSheet = getOrCreateRestockSheet(ss);
  const now = new Date();
  const timestampStr = payload.timestamp || Utilities.formatDate(now, "GMT+7", "yyyy-MM-dd HH:mm:ss");
  const dateStr = payload.date || Utilities.formatDate(now, "GMT+7", "yyyy-MM-dd");

  restockSheet.appendRow([
    "'" + timestampStr,
    "'" + dateStr,
    adminName,
    cfg.title,
    payload.name,
    newStock - prevStock, // ส่วนต่าง
    payload.unit || "อัน",
    prevStock,
    newStock,
    payload.note || "แก้ไขสต็อกคงเหลือโดยตรง"
  ]);

  return {
    success: true,
    message: "แก้ไขยอดสต็อกสำเร็จ",
    prevStock: prevStock,
    newStock: newStock,
    targetRow: targetRow
  };
}

/**
 * ดึงประวัติการเบิกทั้งหมด (ใช้ getDisplayValues() เพื่อรักษาเวลาที่แท้จริง ไม่ให้เพี้ยนเป็น Timezone อื่น)
 */
function getRequisitionLogs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(LOG_SHEET_REQUISITION);
  if (!sheet) return [];
  
  const displayValues = sheet.getDataRange().getDisplayValues();
  if (displayValues.length <= 1) return [];
  
  const logs = [];
  for (let r = 1; r < displayValues.length; r++) {
    const row = displayValues[r];
    if (!row[0] && !row[3]) continue;
    logs.push({
      timestamp: row[0] ? String(row[0]).trim() : "",
      date: row[1] ? String(row[1]).trim() : "",
      time: row[2] ? String(row[2]).trim() : "",
      requester: row[3] ? String(row[3]).trim() : "",
      department: row[4] ? String(row[4]).trim() : "",
      itemName: row[5] ? String(row[5]).trim() : "",
      qty: isNaN(Number(row[6])) ? 0 : Number(row[6]),
      unit: row[7] ? String(row[7]).trim() : "",
      prevStock: isNaN(Number(row[8])) ? 0 : Number(row[8]),
      newStock: isNaN(Number(row[9])) ? 0 : Number(row[9]),
      note: row[10] ? String(row[10]).trim() : ""
    });
  }
  
  return logs.reverse();
}

function getRestockLogs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(LOG_SHEET_RESTOCK);
  if (!sheet) return [];
  
  const displayValues = sheet.getDataRange().getDisplayValues();
  if (displayValues.length <= 1) return [];
  
  const logs = [];
  for (let r = 1; r < displayValues.length; r++) {
    const row = displayValues[r];
    if (!row[0] && !row[2]) continue;
    logs.push({
      timestamp: row[0] ? String(row[0]).trim() : "",
      date: row[1] ? String(row[1]).trim() : "",
      adminName: row[2] ? String(row[2]).trim() : "",
      department: row[3] ? String(row[3]).trim() : "",
      itemName: row[4] ? String(row[4]).trim() : "",
      qty: isNaN(Number(row[5])) ? 0 : Number(row[5]),
      unit: row[6] ? String(row[6]).trim() : "",
      prevStock: isNaN(Number(row[7])) ? 0 : Number(row[7]),
      newStock: isNaN(Number(row[8])) ? 0 : Number(row[8]),
      note: row[9] ? String(row[9]).trim() : ""
    });
  }
  
  return logs.reverse();
}

function getOrCreateLogSheet(ss) {
  let sheet = ss.getSheetByName(LOG_SHEET_REQUISITION);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET_REQUISITION);
    const headers = [
      "วัน-เวลา",
      "วันที่",
      "เวลา",
      "ชื่อผู้เบิก",
      "สาขา/แผนก",
      "รายการวัสดุ",
      "จำนวนเบิก",
      "หน่วย",
      "สต็อกก่อนเบิก",
      "สต็อกคงเหลือ",
      "หมายเหตุ"
    ];
    sheet.appendRow(headers);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground("#1e293b");
    headerRange.setFontColor("#ffffff");
    headerRange.setFontWeight("bold");
    sheet.setFrozenRows(1);
    
    // ตั้งค่าคอลัมน์ A, B, C ให้เป็น Plain Text
    sheet.getRange(2, 1, sheet.getMaxRows() - 1, 3).setNumberFormat("@");
  }
  return sheet;
}

function getOrCreateRestockSheet(ss) {
  let sheet = ss.getSheetByName(LOG_SHEET_RESTOCK);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET_RESTOCK);
    const headers = [
      "วัน-เวลา",
      "วันที่",
      "ผู้เติมสต็อก",
      "สาขา/แผนก",
      "รายการวัสดุ",
      "จำนวนเติม",
      "หน่วย",
      "สต็อกก่อนเติม",
      "สต็อกหลังเติม",
      "หมายเหตุ"
    ];
    sheet.appendRow(headers);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground("#065f46");
    headerRange.setFontColor("#ffffff");
    headerRange.setFontWeight("bold");
    sheet.setFrozenRows(1);
    sheet.getRange(2, 1, sheet.getMaxRows() - 1, 2).setNumberFormat("@");
  }
  return sheet;
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
