/**
 * ชื่อชีตที่ใช้เก็บข้อมูล
 */
const IMPORT_SHEET_NAME = "ข้อมูลขนส่ง";

/**
 * โหลดรายการแจ้งนำเข้าของสมาชิก (ใช้แสดงในตาราง)
 */
function loadImportItems(memberId) {
  try {
    if (!memberId) return [];

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(IMPORT_SHEET_NAME);
    if (!sheet) return [];

    var data = sheet.getDataRange().getValues();
    var result = [];
    var targetMemberId = memberId.toString().trim().toLowerCase();

    for (var i = 1; i < data.length; i++) {
      if (!data[i][0] && !data[i][1]) continue; // ข้ามแถวว่าง

      var rowOrderNum = data[i][0] ? data[i][0].toString().trim() : "";
      var rowMemberId = data[i][1] ? data[i][1].toString().trim().toLowerCase() : "";
      var rowTracking = data[i][2] ? data[i][2].toString().trim() : "";
      var rowDetail   = data[i][3] ? data[i][3].toString().trim() : "";
      var rowWeight   = data[i][5] ? data[i][5].toString().trim() : "";
      var rowCbm      = data[i][6] ? data[i][6].toString().trim() : "";
      
      if (rowMemberId === targetMemberId) {
        result.push({
          orderNum: rowOrderNum, 
          tracking: rowTracking, 
          detail: rowDetail,
          weight: rowWeight,
          cbm: rowCbm
        });
      }
    }
    
    return result.reverse(); 
  } catch (error) {
    throw new Error("ไม่สามารถโหลดข้อมูลได้: " + error.message);
  }
}

/**
 * เพิ่มรายการแจ้งนำเข้าใหม่ (คืนค่า data กลับไปด้วย)
 */
function addItemToSheet(memberId, tracking, detail, weight, cbm) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(IMPORT_SHEET_NAME);
    if (!sheet) return { success: false, message: "ไม่พบชีตข้อมูล" };

    var newOrderNum = generateOrderNumber(sheet);
    var timestamp = Utilities.formatDate(new Date(), "Asia/Bangkok", "dd/MM/yyyy HH:mm:ss");

    var newRow = new Array(18).fill("");
    newRow[0] = newOrderNum;       
    newRow[1] = memberId;          
    newRow[2] = tracking;          
    newRow[3] = detail;            
    newRow[5] = weight;
    newRow[6] = cbm;
    newRow[11] = "รอเข้าโกดังจีน";   
    newRow[12] = timestamp;        

    sheet.appendRow(newRow);
    
    // ส่งข้อมูลที่เพิ่มสำเร็จกลับไปให้หน้าเว็บ เพื่อเอาไปต่อตารางทันที
    return { 
      success: true, 
      message: "เพิ่มรายการเข้าสู่ระบบเรียบร้อยแล้ว",
      data: {
        orderNum: newOrderNum,
        tracking: tracking,
        detail: detail,
        weight: weight,
        cbm: cbm
      }
    };
  } catch (error) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + error.message };
  }
}

/**
 * อัปเดตรายการสินค้า
 */
function updateItemInSheet(orderNum, tracking, detail, weight, cbm) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(IMPORT_SHEET_NAME);
    if (!sheet) return { success: false, message: "ไม่พบชีตข้อมูล" };

    var data = sheet.getDataRange().getValues();
    var foundRow = -1;

    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      if (data[i][0].toString().trim() === orderNum.toString().trim()) {
        foundRow = i + 1; 
        break;
      }
    }

    if (foundRow !== -1) {
      sheet.getRange(foundRow, 3).setValue(tracking); 
      sheet.getRange(foundRow, 4).setValue(detail);   
      sheet.getRange(foundRow, 6).setValue(weight);
      sheet.getRange(foundRow, 7).setValue(cbm);
      return { success: true, message: "แก้ไขข้อมูลเรียบร้อยแล้ว" };
    } else {
      return { success: false, message: "ไม่พบรหัสรายการนี้ในระบบ" };
    }
  } catch (error) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + error.message };
  }
}

/**
 * ลบรายการสินค้า
 */
function deleteItemFromSheet(orderNum) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(IMPORT_SHEET_NAME);
    if (!sheet) return { success: false, message: "ไม่พบชีตข้อมูล" };

    var data = sheet.getDataRange().getValues();
    var foundRow = -1;

    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      if (data[i][0].toString().trim() === orderNum.toString().trim()) {
        foundRow = i + 1;
        break;
      }
    }

    if (foundRow !== -1) {
      sheet.deleteRow(foundRow); 
      return { success: true, message: "ลบรายการเรียบร้อยแล้ว" };
    } else {
      return { success: false, message: "ไม่พบรหัสรายการนี้ในระบบ" };
    }
  } catch (error) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + error.message };
  }
}

/**
 * Helper Function: ใช้รันเลข OrderNumber
 */
function generateOrderNumber(sheet) {
  var data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    return "SC00001";
  }

  var lastOrderNum = "";
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][0] && data[i][0].toString().trim() !== "") {
      lastOrderNum = data[i][0].toString().trim();
      break;
    }
  }

  if (lastOrderNum && lastOrderNum.match(/^SC\d+$/)) {
    var numPart = parseInt(lastOrderNum.substring(2), 10);
    var nextNum = numPart + 1;
    var paddedNum = ("00000" + nextNum).slice(-5);
    return "SC" + paddedNum;
  } else {
    var randomId = new Date().getTime().toString().substring(5);
    return "SC" + randomId;
  }
}