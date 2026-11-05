/**
 * ==============================================================================
 * ระบบจัดการธุรกิจ F&B (เครื่องดื่ม & เบเกอรี่) - Backend Google Apps Script (GAS)
 * ==============================================================================
 * ทำหน้าที่เป็น REST API (JSON) ให้บริการ Frontend (POS & จัดการสต็อก)
 * และจัดการข้อมูล 4 ชีทหลัก:
 * 1. Ingredients   : คลังวัตถุดิบและสต็อกคงเหลือ (พร้อมรูปภาพ Base64/Emoji)
 * 2. Recipes       : เมนูอาหาร/เครื่องดื่ม พร้อมสถานะ (Active, Draft, Inactive)
 * 3. Recipe_Items  : โครงสร้างสูตร (ผูกวัตถุดิบและปริมาณการใช้ต่อแก้ว/ชิ้น)
 * 4. Sales_Log     : ประวัติการขายและต้นทุน
 */

// โครงสร้างหัวตาราง (Headers)
const SHEET_CONFIG = {
  Ingredients: [
    "ID", "Name", "Emoji", "PurchasePrice", "PurchaseQty", 
    "PurchaseUnit", "RecipeUnit", "CostPerUnit", "CurrentStock", "Image"
  ],
  Recipes: [
    "ID", "Name", "Emoji", "SellingPrice", "Status", "Image"
  ],
  Recipe_Items: [
    "RecipeID", "IngredientID", "QuantityUsed"
  ],
  Sales_Log: [
    "Timestamp", "RecipeID", "RecipeName", "QtySold", "TotalRevenue", "TotalCost"
  ]
};

/**
 * --------------------------------------------------------------------------
 * 0. เมนูลัดบน Google Sheets (เปิดชีทมาจะเห็นเมนูนี้ด้านบน สะดวกมาก)
 * --------------------------------------------------------------------------
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("⚙️ F&B เมนูจัดการ")
    .addItem("🧹 ล้างข้อมูลทั้งหมดในตาราง (Clear All Data)", "menuClearData")
    .addItem("⚡ ตรวจสอบ/สร้างหัวตาราง (Init Headers)", "menuInitHeaders")
    .addToUi();
}

function menuClearData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  clearAllSheetData(ss);
  SpreadsheetApp.getUi().alert("ล้างข้อมูลทั้งหมดใน 4 ชีทเรียบร้อยแล้ว! (คงเหลือเฉพาะหัวตารางพร้อมใส่ข้อมูลจริง)");
}

function menuInitHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  initializeSheets(ss);
  SpreadsheetApp.getUi().alert("ตรวจสอบและสร้างหัวตารางเรียบร้อยแล้ว!");
}

/**
 * --------------------------------------------------------------------------
 * 1. doGet: ให้บริการดึงข้อมูลทั้งหมด หรือสั่ง Init/Clear ข้อมูล
 * --------------------------------------------------------------------------
 */
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "getAll";
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // ตัวช่วยสร้างโครงสร้างชีทเปล่าเริ่มต้น (ไม่มี mock data)
    if (action === "init") {
      return initializeSheets(ss);
    }

    // ล้างข้อมูลทั้งหมดในตาราง (คงไว้เฉพาะหัวตาราง)
    if (action === "clearData") {
      return clearAllSheetData(ss);
    }

    // ดึงข้อมูลทั้งหมด 4 ชีทเพื่อส่งให้ Frontend
    const responseData = {
      ingredients: getSheetData(ss, "Ingredients"),
      recipes: getSheetData(ss, "Recipes"),
      recipe_items: getSheetData(ss, "Recipe_Items"),
      sales_log: getSheetData(ss, "Sales_Log", 100) // ดึงประวัติขายล่าสุด 100 รายการ
    };

    return createJsonResponse({ success: true, data: responseData });
  } catch (error) {
    return createJsonResponse({ success: false, error: error.toString() });
  }
}

/**
 * --------------------------------------------------------------------------
 * 2. doPost: รับข้อมูลการทำรายการต่าง ๆ (บันทึกขาย, บันทึกวัตถุดิบ, สูตร)
 * --------------------------------------------------------------------------
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return createJsonResponse({ success: false, error: "ไม่พบข้อมูลที่ส่งเข้ามา (Empty payload)" });
    }

    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 2.1 บันทึกยอดขายจากหน้า POS พร้อมตัดสต็อกอัตโนมัติ
    if (action === "recordSale") {
      return handleRecordSale(ss, payload.data);
    }

    // 2.2 บันทึก/เพิ่ม/อัปเดต วัตถุดิบ
    if (action === "saveIngredient") {
      return handleSaveIngredient(ss, payload.data);
    }

    // 2.3 บันทึก/เพิ่ม/อัปเดต สูตรเมนู (Recipes + Recipe_Items)
    if (action === "saveRecipe") {
      return handleSaveRecipe(ss, payload.data);
    }

    // 2.4 ปรับยอดสต็อกโดยตรง (Manual Stock Adjust / เติมของ)
    if (action === "updateStock") {
      return handleUpdateStock(ss, payload.data);
    }

    // 2.5 ล้างข้อมูลทั้งหมดในดาต้าเบส
    if (action === "clearData") {
      return clearAllSheetData(ss);
    }

    return createJsonResponse({ success: false, error: "ไม่พบ Action ที่ระบุ: " + action });
  } catch (error) {
    return createJsonResponse({ success: false, error: error.toString() });
  }
}

/**
 * --------------------------------------------------------------------------
 * 3. ฟังก์ชันบันทึกการขายและตัดสต็อกอัตโนมัติ (Core Business Logic)
 * --------------------------------------------------------------------------
 */
function handleRecordSale(ss, saleData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // รอคิวสูงสุด 10 วินาที

    const salesSheet = ss.getSheetByName("Sales_Log");
    const ingSheet = ss.getSheetByName("Ingredients");
    const recipeItemsSheet = ss.getSheetByName("Recipe_Items");

    if (!salesSheet || !ingSheet || !recipeItemsSheet) {
      throw new Error("ชีทไม่ครบถ้วน กรุณารันคำสั่งสร้างชีทก่อน");
    }

    const itemsSold = Array.isArray(saleData) ? saleData : [saleData];
    const timestamp = new Date().toLocaleString("th-TH", { timeZone: "Asia/Bangkok" });

    // อ่านข้อมูลสูตรทั้งหมดเพื่อคำนวณการตัดสต็อก
    const recipeItemsData = getSheetData(ss, "Recipe_Items");
    
    // โหลดสต็อกปัจจุบันของวัตถุดิบทั้งหมดเข้า Map เพื่อตัดทีเดียว
    const ingRows = ingSheet.getDataRange().getValues();
    const ingHeaders = ingRows[0];
    const idIdx = ingHeaders.indexOf("ID");
    const stockIdx = ingHeaders.indexOf("CurrentStock");

    if (idIdx === -1 || stockIdx === -1) {
      throw new Error("คอลัมน์ ID หรือ CurrentStock ใน Ingredients ไม่ถูกต้อง");
    }

    const stockMap = {};
    for (let i = 1; i < ingRows.length; i++) {
      const id = String(ingRows[i][idIdx]).trim();
      if (id) {
        stockMap[id] = {
          rowNumber: i + 1,
          stock: parseFloat(ingRows[i][stockIdx]) || 0
        };
      }
    }

    for (const sale of itemsSold) {
      const recipeId = sale.recipeId;
      const recipeName = sale.recipeName || "";
      const qtySold = parseFloat(sale.qtySold) || 1;
      const totalRevenue = parseFloat(sale.totalRevenue) || 0;
      const totalCost = parseFloat(sale.totalCost) || 0;

      // 1. บันทึกลง Sales_Log
      salesSheet.appendRow([
        timestamp,
        recipeId,
        recipeName,
        qtySold,
        totalRevenue,
        totalCost
      ]);

      // 2. คำนวณตัดสต็อกตามสูตร
      let itemsToDeduct = sale.itemsUsed;
      if (!itemsToDeduct || itemsToDeduct.length === 0) {
        itemsToDeduct = recipeItemsData
          .filter(r => String(r.RecipeID).trim() === String(recipeId).trim())
          .map(r => ({
            ingredientId: r.IngredientID,
            qty: parseFloat(r.QuantityUsed) || 0
          }));
      }

      for (const item of itemsToDeduct) {
        const ingId = String(item.ingredientId).trim();
        const usedQty = (parseFloat(item.qty) || 0) * qtySold;

        if (stockMap[ingId]) {
          stockMap[ingId].stock -= usedQty;
        }
      }
    }

    // 3. บันทึกค่าสต็อกใหม่ลงชีท Ingredients
    for (const ingId in stockMap) {
      const item = stockMap[ingId];
      ingSheet.getRange(item.rowNumber, stockIdx + 1).setValue(Math.round(item.stock * 100) / 100);
    }

    return createJsonResponse({
      success: true,
      message: "บันทึกยอดขายและหักสต็อกวัตถุดิบเรียบร้อยแล้ว"
    });
  } catch (err) {
    return createJsonResponse({ success: false, error: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

/**
 * --------------------------------------------------------------------------
 * 4. ฟังก์ชันบันทึกวัตถุดิบ (เพิ่มใหม่ หรือ อัปเดตข้อมูล)
 * --------------------------------------------------------------------------
 */
function handleSaveIngredient(ss, data) {
  const sheet = ss.getSheetByName("Ingredients");
  if (!sheet) throw new Error("ไม่พบชีท Ingredients");

  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const idIdx = headers.indexOf("ID");

  const id = data.id || ("ING_" + Utilities.getUuid().substring(0, 5).toUpperCase());
  const name = data.name || "";
  const emoji = data.emoji || "🍂";
  const purchasePrice = parseFloat(data.purchasePrice) || 0;
  const purchaseQty = parseFloat(data.purchaseQty) || 1;
  const purchaseUnit = data.purchaseUnit || "แพ็ก";
  const recipeUnit = data.recipeUnit || "กรัม";
  const costPerUnit = parseFloat(data.costPerUnit) || (purchasePrice / purchaseQty);
  const currentStock = parseFloat(data.currentStock) || 0;
  const image = data.image || "";

  let existingRow = -1;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][idIdx]).trim() === String(id).trim()) {
      existingRow = i + 1;
      break;
    }
  }

  const rowValues = [
    id, name, emoji, purchasePrice, purchaseQty, 
    purchaseUnit, recipeUnit, costPerUnit, currentStock, image
  ];

  if (existingRow > 0) {
    sheet.getRange(existingRow, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }

  return createJsonResponse({
    success: true,
    message: "บันทึกวัตถุดิบสำเร็จ",
    id: id
  });
}

/**
 * --------------------------------------------------------------------------
 * 5. ฟังก์ชันบันทึกสูตรเมนู (Recipes + Recipe_Items)
 * --------------------------------------------------------------------------
 */
function handleSaveRecipe(ss, data) {
  const recipeSheet = ss.getSheetByName("Recipes");
  const itemsSheet = ss.getSheetByName("Recipe_Items");
  if (!recipeSheet || !itemsSheet) throw new Error("ชีทไม่ครบถ้วน");

  const id = data.id || ("REC_" + Utilities.getUuid().substring(0, 5).toUpperCase());
  const name = data.name || "";
  const emoji = data.emoji || "🍹";
  const sellingPrice = parseFloat(data.sellingPrice) || 0;
  const image = data.image || "";
  let status = data.status || "Draft";
  if (status === "Active" && sellingPrice <= 0) {
    status = "Draft";
  }

  // 1. บันทึกตาราง Recipes
  const rRows = recipeSheet.getDataRange().getValues();
  const rHeaders = rRows[0];
  const rIdIdx = rHeaders.indexOf("ID");

  let existingRecipeRow = -1;
  for (let i = 1; i < rRows.length; i++) {
    if (String(rRows[i][rIdIdx]).trim() === String(id).trim()) {
      existingRecipeRow = i + 1;
      break;
    }
  }

  const recipeValues = [id, name, emoji, sellingPrice, status, image];
  if (existingRecipeRow > 0) {
    recipeSheet.getRange(existingRecipeRow, 1, 1, recipeValues.length).setValues([recipeValues]);
  } else {
    recipeSheet.appendRow(recipeValues);
  }

  // 2. บันทึกตาราง Recipe_Items (ลบของเดิมที่ผูกกับ recipeId นี้ออกก่อนแล้วใส่ใหม่)
  const itemRows = itemsSheet.getDataRange().getValues();
  for (let i = itemRows.length - 1; i >= 1; i--) {
    if (String(itemRows[i][0]).trim() === String(id).trim()) {
      itemsSheet.deleteRow(i + 1);
    }
  }

  if (Array.isArray(data.items) && data.items.length > 0) {
    const newItems = data.items.map(item => [
      id,
      item.ingredientId,
      parseFloat(item.quantityUsed) || 0
    ]);
    itemsSheet.getRange(itemsSheet.getLastRow() + 1, 1, newItems.length, 3).setValues(newItems);
  }

  return createJsonResponse({
    success: true,
    message: "บันทึกสูตรเมนูเรียบร้อย",
    id: id
  });
}

/**
 * --------------------------------------------------------------------------
 * 6. ฟังก์ชันปรับสต็อกคงเหลือโดยตรง (Manual Adjustment / เติมของ)
 * --------------------------------------------------------------------------
 */
function handleUpdateStock(ss, data) {
  const sheet = ss.getSheetByName("Ingredients");
  const rows = sheet.getDataRange().getValues();
  const idIdx = rows[0].indexOf("ID");
  const stockIdx = rows[0].indexOf("CurrentStock");

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][idIdx]).trim() === String(data.id).trim()) {
      const newStock = parseFloat(data.newStock);
      sheet.getRange(i + 1, stockIdx + 1).setValue(newStock);
      return createJsonResponse({ success: true, message: "อัปเดตสต็อกสำเร็จ", newStock: newStock });
    }
  }

  return createJsonResponse({ success: false, error: "ไม่พบรหัสวัตถุดิบ: " + data.id });
}

/**
 * --------------------------------------------------------------------------
 * 7. ฟังก์ชันล้างข้อมูลทั้งหมดในตาราง (Clear Data - คงไว้แค่ Header)
 * --------------------------------------------------------------------------
 */
function clearAllSheetData(ss) {
  const sheetNames = ["Ingredients", "Recipes", "Recipe_Items", "Sales_Log"];
  sheetNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastRow > 1 && lastCol > 0) {
        // ล้างข้อมูลตั้งแต่แถว 2 ลงไป (รักษาหัวตารางแถว 1 ไว้)
        sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
      }
    }
  });

  return createJsonResponse({
    success: true,
    message: "ล้างข้อมูลในดาต้าเบสเรียบร้อยแล้ว (ตารางสะอาด 100% พร้อมใส่ข้อมูลจริง)"
  });
}

/**
 * --------------------------------------------------------------------------
 * Helper Functions & Init Tool
 * --------------------------------------------------------------------------
 */
function getSheetData(ss, sheetName, limit) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];

  const headers = rows[0].map(h => String(h).trim());
  let dataRows = rows.slice(1);

  if (limit && limit > 0) {
    dataRows = dataRows.slice(-limit).reverse();
  }

  return dataRows.map(row => {
    const obj = {};
    headers.forEach((header, i) => {
      obj[header] = row[i];
    });
    return obj;
  });
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * ฟังก์ชันสร้างชีทและหัวตารางเปล่าเริ่มต้น (แบบไม่มีข้อมูลตัวอย่าง)
 */
function initializeSheets(ss) {
  const sheetsToCreate = [
    { name: "Ingredients", headers: SHEET_CONFIG.Ingredients },
    { name: "Recipes", headers: SHEET_CONFIG.Recipes },
    { name: "Recipe_Items", headers: SHEET_CONFIG.Recipe_Items },
    { name: "Sales_Log", headers: SHEET_CONFIG.Sales_Log }
  ];

  sheetsToCreate.forEach(cfg => {
    let sheet = ss.getSheetByName(cfg.name);
    if (!sheet) {
      sheet = ss.insertSheet(cfg.name);
    }
    // หากชีทว่าง ให้ใส่เฉพาะหัวตาราง
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(cfg.headers);
      const headerRange = sheet.getRange(1, 1, 1, cfg.headers.length);
      headerRange.setBackground("#1e293b").setFontColor("#ffffff").setFontWeight("bold");
      sheet.setFrozenRows(1);
    } else {
      // ตรวจสอบและอัปเดตหัวตารางแถวแรกให้ตรงกับ Config ล่าสุด
      sheet.getRange(1, 1, 1, cfg.headers.length).setValues([cfg.headers]);
    }
  });

  // ลบแผ่นงานเริ่มต้นที่เป็นแผ่นงานเปล่า (Sheet1 หรือ แผ่นงาน1) ออกอัตโนมัติ
  try {
    const defaultSheet = ss.getSheetByName("Sheet1") || ss.getSheetByName("แผ่นงาน1");
    if (defaultSheet && defaultSheet.getLastRow() === 0 && ss.getSheets().length > 1) {
      ss.deleteSheet(defaultSheet);
    }
  } catch (e) {
    // ข้ามกรณีลบไม่ได้
  }

  return createJsonResponse({
    success: true,
    message: "สร้างโครงสร้างชีทเปล่าเรียบร้อยแล้ว! พร้อมใช้งาน 100%"
  });
}
