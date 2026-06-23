function updateUserProfile(data) {
  try {
    var sheet = getSheetByName("สมาชิก");
    if (!sheet) {
      return { success: false, message: "ไม่พบชีตชื่อ 'สมาชิก' ในฐานข้อมูล" };
    }
    
    var dataRange = sheet.getDataRange().getValues();
    var memberRow = -1;
    
    // 1. ค้นหาบรรทัดของสมาชิกคนนี้จาก memberId
    for (var i = 1; i < dataRange.length; i++) {
      if (dataRange[i][0].toString().trim() === data.memberId) {
        memberRow = i + 1; 
        break;
      }
    }
    
    if (memberRow === -1) {
      return { success: false, message: "ไม่พบข้อมูลรหัสสมาชิกนี้ในระบบ" };
    }

    // 2. เช็คข้อมูลซ้ำในระบบ (ยกเว้นบรรทัดของตัวเอง)
    var inputPhone = data.phone.trim();
    var inputFullname = data.fullname.trim();

    for (var i = 1; i < dataRange.length; i++) {
      if ((i + 1) !== memberRow) { 
        var existingFullname = dataRange[i][1].toString().trim();
        var existingPhone = dataRange[i][2].toString().trim();
        
        // เช็คเบอร์โทรศัพท์
        if (existingPhone === inputPhone) { 
          return { success: false, message: "เบอร์โทรศัพท์นี้ถูกใช้งานโดยบัญชีอื่นแล้ว" };
        }
        
        // เช็คชื่อ-นามสกุล
        if (existingFullname === inputFullname) {
          return { success: false, message: "ชื่อ-นามสกุลนี้มีอยู่ในระบบแล้ว โปรดตรวจสอบอีกครั้ง" };
        }
      }
    }
    
    // 3. ทำการอัปเดตข้อมูล (คอลัมน์ B=ชื่อ, C=เบอร์, D=ที่อยู่)
    sheet.getRange(memberRow, 2).setValue(data.fullname);
    sheet.getRange(memberRow, 3).setValue("'" + data.phone); 
    sheet.getRange(memberRow, 4).setValue(data.address);
    
    // 4. ถ้ามีการกรอกรหัสผ่านใหม่มา ให้ทำการอัปเดต (คอลัมน์ F)
    if (data.password && data.password.trim() !== "") {
      var passwordHash = Utilities.base64Encode(
        Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, data.password)
      );
      sheet.getRange(memberRow, 6).setValue(passwordHash);
    }
    
    // 5. คืนค่าข้อมูลอัปเดตกลับไปให้ Session ฝั่ง Client
    var updatedUserData = {
      memberId: data.memberId,
      fullname: data.fullname,
      phone: data.phone,
      address: data.address,
      role: data.role 
    };
    
    return { 
      success: true, 
      message: "อัปเดตข้อมูลสำเร็จ",
      user: updatedUserData
    };
    
  } catch (error) {
    return { success: false, message: "เกิดข้อผิดพลาดในการบันทึกข้อมูล: " + error.message };
  }
}