/**
 * ตั้งชื่อชีตเป็น var เพื่อป้องกันปัญหาตัวแปรซ้ำกับไฟล์อื่น (เช่น user_importBE.gs)
 */
var SETTING_SHEET_NAME_PAYMENT = "ตั้งค่า";
var BILL_SHEET_NAME_PAYMENT = "บิลชำระเงิน";
var IMPORT_SHEET_NAME_PAYMENT = "ข้อมูลขนส่ง";
var RATE_SHEET_NAME_PAYMENT = "อัตราค่าขนส่ง"; // เปลี่ยนชื่อตัวแปรเพื่อป้องกันการซ้ำซ้อน

/**
 * ฟังก์ชันช่วยดึงเฉพาะตัวเลขจากข้อความ (เพื่อเอาไปคำนวณ)
 */
function extractNumberForCalculate(str) {
  if (!str) return 0;
  var match = str.toString().match(/\d+(\.\d+)?/);
  return match ? parseFloat(match[0]) : 0;
}

/**
 * ดึงข้อมูลพัสดุและตั้งค่าเพื่อนำมาแสดงในหน้าชำระเงิน
 */
function getPaymentPageData(memberId) {
  try {
    if (!memberId) throw new Error("ไม่พบข้อมูลสมาชิก");

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var settingSheet = ss.getSheetByName(SETTING_SHEET_NAME_PAYMENT);
    var transportSheet = ss.getSheetByName(IMPORT_SHEET_NAME_PAYMENT);
    var rateSheet = ss.getSheetByName(RATE_SHEET_NAME_PAYMENT); // อ้างอิงตัวแปรที่แก้ชื่อแล้ว

    if (!settingSheet || !transportSheet || !rateSheet) {
      throw new Error("ไม่พบชีตฐานข้อมูล");
    }

    // 1. ดึงข้อมูลการตั้งค่า (ค่าบริการพื้นฐาน และ ธนาคาร)
    var settingsData = settingSheet.getDataRange().getValues();
    var settings = { baseFee: "", bankName: "", bankAcc: "", bankAccName: "" };

    if (settingsData.length > 1) {
      settings = {
        baseFee: settingsData[1][0] || "",             
        bankName: settingsData[1][5] || "",            
        bankAcc: settingsData[1][6] || "",             
        bankAccName: settingsData[1][7] || ""          
      };
    }

    // 2. ดึงข้อมูลอัตราค่าขนส่งจากชีต "อัตราค่าขนส่ง"
    var rateData = rateSheet.getDataRange().getValues();
    var rates = [];
    for (var r = 1; r < rateData.length; r++) {
      if (rateData[r][0]) {
        rates.push({
          transportType: rateData[r][0].toString().trim(),
          productType: rateData[r][1] ? rateData[r][1].toString().trim() : "",
          days: rateData[r][2] ? rateData[r][2].toString().trim() : "",
          priceKg: parseFloat(rateData[r][3]) || 0,
          priceCbm: parseFloat(rateData[r][4]) || 0
        });
      }
    }

    // 3. ดึงข้อมูลพัสดุของลูกค้าคนนี้ที่ยังไม่มีเลขบิล
    var transportData = transportSheet.getDataRange().getValues();
    var parcels = [];
    var targetMemberId = memberId.toString().trim().toLowerCase();

    for (var i = 1; i < transportData.length; i++) {
      var rowMemberId = transportData[i][1] ? transportData[i][1].toString().trim().toLowerCase() : "";
      var rowBillId = transportData[i][18] ? transportData[i][18].toString().trim() : ""; // คอลัมน์ S (Index 18): Bill_ID
      
      // ดึงเฉพาะของ member นี้ และ ต้องยังไม่มีบิลชำระเงิน
      if (rowMemberId === targetMemberId && rowBillId === "") { 
        var weightStr = transportData[i][6] ? transportData[i][6].toString().trim() : "0"; // คอลัมน์ G (Index 6)
        var cbmStr = transportData[i][7] ? transportData[i][7].toString().trim() : "0";    // คอลัมน์ H (Index 7)
        
        parcels.push({
          orderNum: transportData[i][0] ? transportData[i][0].toString() : "", 
          tracking: transportData[i][2] ? transportData[i][2].toString() : "", 
          weight: parseFloat(weightStr) || 0,
          cbm: parseFloat(cbmStr) || 0
        });
      }
    }

    return { success: true, settings: settings, rates: rates, parcels: parcels };

  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * บันทึกข้อมูลการชำระเงิน
 */
function submitPaymentData(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var billSheet = ss.getSheetByName(BILL_SHEET_NAME_PAYMENT); // อ้างอิงตัวแปรที่แก้ชื่อแล้ว
    var importSheet = ss.getSheetByName(IMPORT_SHEET_NAME_PAYMENT);

    if (!billSheet || !importSheet) throw new Error("ไม่พบชีตสำหรับบันทึกข้อมูล");

    // 1. แปลงไฟล์ภาพสลิปจาก Base64 และสร้างลง Google Drive
    var slipUrl = "";
    if (payload.slipBase64) {
      var blob = Utilities.newBlob(Utilities.base64Decode(payload.slipBase64), payload.slipMime, payload.slipName);
      var folder = DriveApp.getFolderById("1bJHkMVxMdXPDJ3XRZJzcTRfrNYYXLbLm"); // เปลี่ยน Folder ID ของคุณถ้าต้องการ
      var file = folder.createFile(blob);
      try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch(e) {}
      slipUrl = file.getUrl();
    }

    // 2. สร้าง Bill_ID (เช่น INV20260625010101)
    var timestampStr = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyyMMddHHmmss");
    var billId = "INV" + timestampStr;
    var currentTime = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss");

    // 3. บันทึกลงชีต "บิลชำระเงิน"
    billSheet.appendRow([
      billId,                           // A: Bill_ID
      payload.memberId,                 // B: รหัสสมาชิก
      payload.orderList.join(", "),     // C: Order_List
      payload.totalTransportFee,        // D: ยอดรวมขนส่งจีน-ไทย
      "",                               // E: วิธีจัดส่งในไทย
      "",                               // F: ค่าจัดส่งในไทย
      payload.finalTotal,               // G: ยอดสุทธิ
      slipUrl,                          // H: สลิปโอนเงิน_URL
      "รอตรวจสอบยอด",                     // I: สถานะบิล
      "",                               // J: Tracking_Thai
      currentTime                       // K: วันที่ทำรายการชำระเงิน
    ]);

    // 4. อัปเดตข้อมูลในชีต "ข้อมูลขนส่ง"
    var importData = importSheet.getDataRange().getValues();
    var targetOrders = payload.orderList;

    for (var i = 1; i < importData.length; i++) {
      var rowOrderNum = importData[i][0] ? importData[i][0].toString().trim() : "";
      
      // ถ้ารายการนี้อยู่ในรายการที่ลูกค้าเลือกชำระเงิน
      if (targetOrders.indexOf(rowOrderNum) !== -1) {
        importSheet.getRange(i + 1, 5).setValue(payload.shippingType);  // คอลัมน์ E (5): ประเภทขนส่ง (EK / SEA)
        importSheet.getRange(i + 1, 6).setValue(payload.productType);   // คอลัมน์ F (6): ประเภทสินค้า
        importSheet.getRange(i + 1, 19).setValue(billId);               // คอลัมน์ S (19): Bill_ID
      }
    }

    return { success: true, message: "บันทึกข้อมูลและอัปโหลดสลิปเรียบร้อยแล้ว!" };

  } catch (error) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + error.message };
  }
}