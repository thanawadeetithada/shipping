/**
 * ดึงข้อมูลอัตราค่าขนส่งจากชีต "อัตราค่าขนส่ง"
 */
function getPricingData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("อัตราค่าขนส่ง");
    
    if (!sheet) {
      throw new Error("ไม่พบชีต 'อัตราค่าขนส่ง' กรุณาสร้างชีตและใส่ข้อมูลให้ถูกต้อง");
    }

    var data = sheet.getDataRange().getValues();
    var pricingMap = {};
    var result = [];

    // เริ่มวนลูปจากแถวที่ 2 (index 1) โดยข้ามหัวตาราง
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var transportType = row[0] ? row[0].toString().trim() : ""; // คอลัมน์ A
      
      if (!transportType) continue;

      var productType = row[1] ? row[1].toString().trim() : ""; // คอลัมน์ B
      var duration = row[2] ? row[2].toString().trim() : "";    // คอลัมน์ C
      var priceKg = row[3] ? row[3].toString().trim() : "-";    // คอลัมน์ D
      var priceCbm = row[4] ? row[4].toString().trim() : "-";   // คอลัมน์ E

      // ถ้ายังไม่มีหมวดหมู่ขนส่งนี้ ให้สร้างใหม่
      if (!pricingMap[transportType]) {
        pricingMap[transportType] = {
          type: transportType,
          time: duration,
          items: []
        };
        result.push(pricingMap[transportType]);
      }

      // นำประเภทสินค้าไปใส่ในหมวดหมู่นั้นๆ
      if (productType) {
        // อัปเดตระยะเวลาให้เป็นข้อมูลล่าสุดของหมวดหมู่นั้นเสมอ
        if(duration) pricingMap[transportType].time = duration; 
        
        // จัดรูปแบบราคาให้มีลูกน้ำ (คอมมา)
        var formatPrice = function(val) {
           if (isNaN(val) || val === "-") return val;
           return Number(val).toLocaleString();
        };

        pricingMap[transportType].items.push({
          name: productType,
          kg: formatPrice(priceKg),
          cbm: formatPrice(priceCbm)
        });
      }
    }

    return { success: true, data: result };

  } catch (error) {
    return { success: false, message: error.message };
  }
}