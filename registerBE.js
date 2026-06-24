function getSheetByName(sheetName) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
}

function registerNewUser(data) {
  try {
    var sheet = getSheetByName("สมาชิก");
    if (!sheet) {
      return { success: false, message: "ไม่พบชีตชื่อ 'สมาชิก' ในฐานข้อมูล" };
    }
    var dataRange = sheet.getDataRange().getValues();
    var newName = data.fullname ? data.fullname.toString().trim() : "";
    
    for (var i = 1; i < dataRange.length; i++) {
      var existingName = dataRange[i][1] ? dataRange[i][1].toString().trim() : "";
      if (existingName === newName && newName !== "") { 
        return { success: false, message: "ชื่อ-นามสกุลนี้ถูกใช้งานสมัครสมาชิกแล้ว" };
      }
      
      if (dataRange[i][2] == data.phone) { 
        return { success: false, message: "เบอร์โทรศัพท์นี้ถูกใช้งานสมัครสมาชิกแล้ว" };
      }
    }
    
    var lastRow = sheet.getLastRow();
    var newMemberId = "SC89001";
    if (lastRow > 1) { 
      var lastId = sheet.getRange(lastRow, 1).getValue().toString().trim(); 
      if (lastId.startsWith("SC89")) {
        var numPart = parseInt(lastId.replace("SC89", ""), 10);
        if (!isNaN(numPart)) {
          var nextNum = numPart + 1;
          var paddedNum = nextNum.toString().padStart(3, '0');
          newMemberId = "SC89" + paddedNum;
        }
      }
    }
    
    var passwordHash = Utilities.base64Encode(
      Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, data.password)
    );
    
    // ตอนบันทึกข้อมูล ให้ใช้ตัวแปร newName ที่ทำการ trim เรียบร้อยแล้วลงไปด้วยครับ
    var rowData = [
      newMemberId,
      newName, 
      "'" + data.phone,
      data.address,
      'user',
      passwordHash
    ];
    sheet.appendRow(rowData);
    
    return { 
      success: true, 
      memberId: newMemberId,
      message: "สมัครสมาชิกสำเร็จ" 
    };
    
  } catch (error) {
    return { success: false, message: "เกิดข้อผิดพลาดในการบันทึกข้อมูล: " + error.message };
  }
}