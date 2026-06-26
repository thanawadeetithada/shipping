function verifyLogin(username, password) {
  try {
    var sheet = getSheetByName("สมาชิก"); 
    if (!sheet) {
      return { success: false, message: "ไม่พบชีตชื่อ 'สมาชิก' ในระบบ" };
    }

    var dataRange = sheet.getDataRange().getValues();
    var inputPasswordHash = Utilities.base64Encode(
      Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password)
    );

    for (var i = 1; i < dataRange.length; i++) {
      var storedFullname = dataRange[i][1].toString().trim();
      var storedPhone = dataRange[i][2].toString().trim();
      var storedHash = dataRange[i][5].toString().trim();

      if (storedPhone === username || storedFullname === username) {
        
        if (storedHash === inputPasswordHash) {
          var userRole = dataRange[i][4].toString().trim().toLowerCase();
          
          var userData = {
            memberId: dataRange[i][0].toString().trim(),
            fullname: storedFullname,                    
            phone: storedPhone,                          
            address: dataRange[i][3].toString().trim(),  
            role: userRole                               
          };
          
          var targetPage = "user_dashboard";
          
          if (userRole === "admin") {
            targetPage = "admin_dashboard";
          } else if (userRole === "user") {
            targetPage = "user_dashboard";
          }
          
          return { 
             success: true, 
             message: "เข้าสู่ระบบสำเร็จ",
             user: userData,
             redirectPage: targetPage
          };
          
        } else {
          return { success: false, message: "เบอร์โทรศัพท์/ชื่อผู้ใช้ หรือ รหัสผ่านไม่ถูกต้อง" };
        }
      }
    }

    return { success: false, message: "เบอร์โทรศัพท์/ชื่อผู้ใช้ หรือ รหัสผ่านไม่ถูกต้อง" };

  } catch (error) {
    return { success: false, message: "เกิดข้อผิดพลาดของระบบ: " + error.message };
  }
}