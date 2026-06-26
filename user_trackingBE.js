var TRACK_SHEET_NAME = "ข้อมูลขนส่ง";
function getTrackingData(memberId) {
  try {
    if (!memberId) throw new Error("ไม่พบข้อมูลสมาชิก");
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(TRACK_SHEET_NAME);
    if (!sheet) throw new Error("ไม่พบชีตฐานข้อมูล 'ข้อมูลขนส่ง'");

    var data = sheet.getDataRange().getValues();
    var targetMemberId = memberId.toString().trim().toLowerCase();
    var results = [];

    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      
      var rowMemberId = data[i][1] ? data[i][1].toString().trim().toLowerCase() : "";
      
      if (rowMemberId === targetMemberId) {
        
         var dateVal = data[i][12]; 
         if(data[i][13]) dateVal = data[i][13];
         if(data[i][14]) dateVal = data[i][14];
         if(data[i][15]) dateVal = data[i][15];
         if(data[i][16]) dateVal = data[i][16];
         
         var formattedDate = "-";
         if (dateVal) {
           var d = new Date(dateVal);
           if (!isNaN(d.getTime())) {
             var months = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];
             formattedDate = d.getDate() + " " + months[d.getMonth()] + " " + (d.getFullYear() + 543); 
           }
         }

         results.push({
             orderNum: data[i][0] ? data[i][0].toString() : "-",
             tracking: data[i][2] ? data[i][2].toString() : "-",
             detail: data[i][3] ? data[i][3].toString() : "-", 
             shippingType: data[i][4] ? data[i][4].toString() : "-",
             weight: data[i][5] ? data[i][5].toString() : "0",
             cbm: data[i][6] ? data[i][6].toString() : "0",
             status: data[i][11] ? data[i][11].toString().trim() : "รอเข้าโกดังจีน", 
             updateDate: formattedDate
         });
      }
    }

    return { success: true, data: results.reverse() }; 
    
  } catch (error) {
    return { success: false, message: error.message };
  }
}