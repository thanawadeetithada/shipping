function getContactData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("ติดต่อเรา");
    
    if (!sheet) throw new Error("ไม่พบชีต 'ติดต่อเรา'");

    var data = sheet.getDataRange().getValues();
    
    if (data.length < 2) throw new Error("ไม่มีข้อมูลการติดต่อ");

    var row = data[1];

    return {
      success: true,
      phone: row[0] ? row[0].toString().trim() : "-",
      line: row[1] ? row[1].toString().trim() : "-",
      lineLink: row[2] ? row[2].toString().trim() : "#",
      email: row[3] ? row[3].toString().trim() : "-"
    };

  } catch (error) {
    return { success: false, message: error.message };
  }
}