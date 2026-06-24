// โหลดข้อมูลของลูกค้าคนนั้นมาแสดงในตาราง (เอาเฉพาะที่สถานะเป็น "รอเข้าโกดังจีน")
function loadImportItems(memberId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ข้อมูลขนส่ง');
  const data = sheet.getDataRange().getValues();
  const result = [];
  
  // ข้ามหัวตาราง เริ่มที่ i=1
  for(let i = 1; i < data.length; i++) {
    if(data[i][1] === memberId && data[i][11] === "รอเข้าโกดังจีน") {
      result.push({
        orderNum: data[i][0], // คอลัมน์ A
        tracking: data[i][2], // คอลัมน์ C
        detail: data[i][3]    // คอลัมน์ D
      });
    }
  }
  return result;
}

// 1. เพิ่มข้อมูล
function addItemToSheet(memberId, tracking, detail) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ข้อมูลขนส่ง');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  
  try {
    const data = sheet.getDataRange().getValues();
    let lastOrderNum = 0;
    for (let i = data.length - 1; i > 0; i--) {
      let orderStr = data[i][0];
      if (orderStr && orderStr.toString().startsWith('SC')) {
        let num = parseInt(orderStr.replace('SC', ''), 10);
        if (!isNaN(num)) { lastOrderNum = num; break; }
      }
    }
    
    let nextOrderStr = "SI-SC" + String(lastOrderNum + 1).padStart(5, '0');
    const timestamp = new Date();
    
    sheet.appendRow([
      nextOrderStr, memberId, tracking, detail, "", "", "", "", "", "", "", 
      "รอเข้าโกดังจีน", timestamp, "", "", "", "", ""
    ]);
    return { success: true, message: 'เพิ่มรายการสำเร็จ' };
  } catch(e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// 2. แก้ไขข้อมูล (อ้างอิงจาก OrderNumber)
function updateItemInSheet(orderNum, newTracking, newDetail) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ข้อมูลขนส่ง');
  const data = sheet.getDataRange().getValues();
  
  for(let i = 1; i < data.length; i++) {
    if(data[i][0] === orderNum) {
      sheet.getRange(i+1, 3).setValue(newTracking); // อัปเดต Tracking
      sheet.getRange(i+1, 4).setValue(newDetail);   // อัปเดต Detail
      return { success: true, message: 'แก้ไขข้อมูลสำเร็จ' };
    }
  }
  return { success: false, message: 'ไม่พบรายการที่ต้องการแก้ไข' };
}

// 3. ลบข้อมูล (อ้างอิงจาก OrderNumber)
function deleteItemFromSheet(orderNum) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ข้อมูลขนส่ง');
  const data = sheet.getDataRange().getValues();
  
  for(let i = 1; i < data.length; i++) {
    if(data[i][0] === orderNum) {
      sheet.deleteRow(i+1);
      return { success: true, message: 'ลบรายการสำเร็จ' };
    }
  }
  return { success: false, message: 'ไม่พบรายการที่ต้องการลบ' };
}