/**
 * ชื่อชีตที่ใช้เก็บข้อมูลพัสดุ
 */
var TRACK_SHEET_NAME = "ข้อมูลขนส่ง";

/**
 * ดึงข้อมูลพัสดุทั้งหมดของสมาชิกคนนั้น เพื่อไปแสดงในหน้าสถานะพัสดุ
 */
function getTrackingData(memberId) {
  try {
    if (!memberId) throw new Error("ไม่พบข้อมูลสมาชิก");
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(TRACK_SHEET_NAME);
    if (!sheet) throw new Error("ไม่พบชีตฐานข้อมูล 'ข้อมูลขนส่ง'");

    var data = sheet.getDataRange().getValues();
    var targetMemberId = memberId.toString().trim().toLowerCase();
    var results = [];

    // วนลูปอ่านข้อมูลทีละบรรทัด
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue; // ข้ามแถวที่ไม่มีเลข OrderNum
      
      var rowMemberId = data[i][1] ? data[i][1].toString().trim().toLowerCase() : "";
      
      // ดึงเฉพาะพัสดุของลูกค้าคนนี้
      if (rowMemberId === targetMemberId) {
        
         // หา 'วันที่อัปเดตล่าสุด' โดยไล่เช็คจากคอลัมน์ Q ถอยกลับไปถึง M (สถานะไหนล่าสุดเอาอันนั้น)
         var dateVal = data[i][12]; // ค่าเริ่มต้น: วันที่แจ้งนำเข้า (M)
         if(data[i][13]) dateVal = data[i][13]; // วันที่ถึงโกดังจีน (N)
         if(data[i][14]) dateVal = data[i][14]; // วันที่กำลังขนส่งมาไทย (O)
         if(data[i][15]) dateVal = data[i][15]; // วันที่ถึงโกดังไทย (P)
         if(data[i][16]) dateVal = data[i][16]; // วันที่จัดส่งสำเร็จ (Q)
         
         // แปลงรูปแบบวันที่ให้สวยงาม
         var formattedDate = "-";
         if (dateVal) {
           var d = new Date(dateVal);
           if (!isNaN(d.getTime())) {
             var months = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];
             formattedDate = d.getDate() + " " + months[d.getMonth()] + " " + (d.getFullYear() + 543); // แปลงเป็น พ.ศ.
           }
         }

         results.push({
             orderNum: data[i][0] ? data[i][0].toString() : "-",
             tracking: data[i][2] ? data[i][2].toString() : "-",
             detail: data[i][3] ? data[i][3].toString() : "-", // คอลัมน์ D: รายละเอียดสินค้า 
             shippingType: data[i][4] ? data[i][4].toString() : "-",
             weight: data[i][5] ? data[i][5].toString() : "0",
             cbm: data[i][6] ? data[i][6].toString() : "0",
             status: data[i][11] ? data[i][11].toString().trim() : "รอเข้าโกดังจีน", // คอลัมน์ L: สถานะสินค้า
             updateDate: formattedDate
         });
      }
    }
    
    // Reverse เพื่อให้ข้อมูลพัสดุที่ใหม่ที่สุดอยู่บนสุด
    return { success: true, data: results.reverse() }; 
    
  } catch (error) {
    return { success: false, message: error.message };
  }
}