/**
 * ตั้งชื่อชีตเป็น var เพื่อป้องกันปัญหาตัวแปรซ้ำกับไฟล์ user_importBE.gs
 */
var SETTING_SHEET_NAME_PAYMENT = "ตั้งค่า";
var BILL_SHEET_NAME = "บิลชำระเงิน";
var IMPORT_SHEET_NAME_PAYMENT = "ข้อมูลขนส่ง";

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

    if (!settingSheet || !transportSheet) {
      throw new Error("ไม่พบชีตฐานข้อมูล");
    }

    // 1. ดึงข้อมูลการตั้งค่า (แถวที่ 2)
    var settingsData = settingSheet.getDataRange().getValues();
    var settings = {
      baseFee: "", carRateText: "", carRateNum: 0, carDays: "", 
      boatRateText: "", boatRateNum: 0, boatDays: "",
      bankName: "", bankAcc: "", bankAccName: ""
    };

    if (settingsData.length > 1) {
      var rawCarRate = settingsData[1][1] ? settingsData[1][1].toString() : "";
      var rawBoatRate = settingsData[1][3] ? settingsData[1][3].toString() : "";

      settings = {
        baseFee: settingsData[1][0] || "",             
        carRateText: rawCarRate,                       
        carRateNum: extractNumberForCalculate(rawCarRate), 
        carDays: settingsData[1][2] ? settingsData[1][2].toString() : "", 
        boatRateText: rawBoatRate,                     
        boatRateNum: extractNumberForCalculate(rawBoatRate), 
        boatDays: settingsData[1][4] ? settingsData[1][4].toString() : "", 
        bankName: settingsData[1][5] || "",            
        bankAcc: settingsData[1][6] || "",             
        bankAccName: settingsData[1][7] || ""          
      };
    }

    // 2. ดึงข้อมูลพัสดุของลูกค้าคนนี้ที่ยังไม่มีเลขบิล
    var transportData = transportSheet.getDataRange().getValues();
    var parcels = [];
    var targetMemberId = memberId.toString().trim().toLowerCase();

    for (var i = 1; i < transportData.length; i++) {
      var rowMemberId = transportData[i][1] ? transportData[i][1].toString().trim().toLowerCase() : "";
      var rowBillId = transportData[i][17] ? transportData[i][17].toString().trim() : ""; // คอลัมน์ R (Index 17): Bill_ID
      
      // ดึงเฉพาะของ member นี้ และ **ต้องยังไม่มีบิลชำระเงิน**
      if (rowMemberId === targetMemberId && rowBillId === "") { 
        var weightStr = transportData[i][5] ? transportData[i][5].toString().trim() : "0";
        var weight = parseFloat(weightStr) || 0;

        parcels.push({
          orderNum: transportData[i][0] ? transportData[i][0].toString() : "", 
          tracking: transportData[i][2] ? transportData[i][2].toString() : "", 
          weight: weight 
        });
      }
    }

    return { success: true, settings: settings, parcels: parcels };

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
    var billSheet = ss.getSheetByName(BILL_SHEET_NAME);
    var importSheet = ss.getSheetByName(IMPORT_SHEET_NAME_PAYMENT);

    if (!billSheet || !importSheet) throw new Error("ไม่พบชีตสำหรับบันทึกข้อมูล");

    // 1. แปลงไฟล์ภาพสลิปจาก Base64 และสร้างลง Google Drive
    var slipUrl = "";
    if (payload.slipBase64) {
      var blob = Utilities.newBlob(Utilities.base64Decode(payload.slipBase64), payload.slipMime, payload.slipName);
      var folder = DriveApp.getFolderById("1bJHkMVxMdXPDJ3XRZJzcTRfrNYYXLbLm");
      var file = folder.createFile(blob);
      
      // ใช้ Try-Catch ครอบไว้ ถ้า Google Block การตั้งค่าแชร์ ก็ให้ข้ามไปเลย โค้ดจะได้ไม่พัง
      try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch(e) {
        // ข้าม Error นี้ไป
      }
      
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
      payload.totalTransportFee,        // D: ยอดรวมขนส่งจีน-ไทย (ค่าขนส่งอย่างเดียว)
      "",                               // E: วิธีจัดส่งในไทย (ปล่อยว่างไว้ก่อน หรือใส่ตามที่รับมา)
      "",                               // F: ค่าจัดส่งในไทย (ปล่อยว่างไว้ก่อน)
      payload.finalTotal,               // G: ยอดสุทธิ
      slipUrl,                          // H: สลิปโอนเงิน_URL
      "รอตรวจสอบยอด",                     // I: สถานะบิล
      "",                               // J: Tracking_Thai
      currentTime                       // K: วันที่ทำรายการชำระเงิน
    ]);

    // 4. อัปเดตข้อมูลในชีต "ข้อมูลขนส่ง" (ประเภทขนส่ง และ Bill_ID)
    var importData = importSheet.getDataRange().getValues();
    var targetOrders = payload.orderList;

    for (var i = 1; i < importData.length; i++) {
      var rowOrderNum = importData[i][0] ? importData[i][0].toString().trim() : "";
      
      // ถ้ารายการนี้อยู่ในรายการที่ลูกค้าเลือกชำระเงิน
      if (targetOrders.indexOf(rowOrderNum) !== -1) {
        importSheet.getRange(i + 1, 5).setValue(payload.shippingType); // คอลัมน์ E (Index 5): ประเภทขนส่ง (EK / SEA)
        importSheet.getRange(i + 1, 18).setValue(billId);              // คอลัมน์ R (Index 18): Bill_ID
      }
    }

    return { success: true, message: "บันทึกข้อมูลและอัปโหลดสลิปเรียบร้อยแล้ว!" };

  } catch (error) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + error.message };
  }
}