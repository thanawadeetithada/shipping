function updateUserProfile(userData) {
  try {
    var sheet;
    if (typeof getSheetByName === 'function') {
      sheet = getSheetByName("สมาชิก");
    } else {
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("สมาชิก");
    }
    
    if (!sheet) {
      return { success: false, message: "ไม่พบข้อมูลชีตชื่อ 'สมาชิก' ในระบบ" };
    }

    var dataRange = sheet.getDataRange().getValues();
    
    var memberId = userData.memberId.toString().trim();
    var inputFullname = userData.fullname.toString().trim();
    var inputPhone = userData.phone.toString().trim();
    
    var foundRow = -1;

    for (var i = 1; i < dataRange.length; i++) {
      if (dataRange[i][0].toString().trim() === memberId) {
        foundRow = i + 1; 
        break;
      }
    }

    if (foundRow === -1) {
      return { success: false, message: "ไม่พบรหัสสมาชิกนี้ในฐานข้อมูลระบบ" };
    }

    for (var j = 1; j < dataRange.length; j++) {
      var currentRow = j + 1;
      
      if (currentRow === foundRow) continue;

      var existingFullname = dataRange[j][1].toString().trim(); 
      var existingPhone = dataRange[j][2].toString().trim();    

      if (existingFullname.toLowerCase() === inputFullname.toLowerCase()) {
        return { success: false, message: "ไม่สามารถบันทึกได้ เนื่องจากชื่อ-นามสกุลนี้ถูกใช้งานแล้วในระบบ" };
      }
      
      if (existingPhone === inputPhone) {
        return { success: false, message: "ไม่สามารถบันทึกได้ เนื่องจากเบอร์โทรศัพท์นี้ถูกใช้งานแล้วในระบบ" };
      }
    }

    sheet.getRange(foundRow, 2).setValue(inputFullname); 
    sheet.getRange(foundRow, 3).setValue(inputPhone);    
    sheet.getRange(foundRow, 4).setValue(userData.address.trim());  

    if (userData.password && userData.password.trim() !== "") {
      var inputPasswordHash = Utilities.base64Encode(
        Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, userData.password.trim())
      );
      sheet.getRange(foundRow, 6).setValue(inputPasswordHash);
    }

    return { success: true, message: "อัปเดตข้อมูลส่วนตัวเสร็จเรียบร้อยแล้ว!" };

  } catch (error) {
    return { success: false, message: "เกิดข้อผิดพลาดของระบบ: " + error.message };
  }
}