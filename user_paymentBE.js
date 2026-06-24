/**
 * ชื่อชีตที่ใช้เก็บข้อมูล
 */
const SETTING_SHEET_NAME = "ตั้งค่า";
// ไม่ต้องประกาศ IMPORT_SHEET_NAME ซ้ำ เพราะดึงมาจาก user_importBE.gs ได้เลย

/**
 * ดึงข้อมูลพัสดุและตั้งค่าเพื่อนำมาแสดงในหน้าชำระเงิน
 */
function getPaymentPageData(memberId) {
  try {
    if (!memberId) throw new Error("ไม่พบข้อมูลสมาชิก");

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var settingSheet = ss.getSheetByName(SETTING_SHEET_NAME);
    var transportSheet = ss.getSheetByName(IMPORT_SHEET_NAME); // ใช้ตัวแปรจากไฟล์อื่นได้เลย

    if (!settingSheet || !transportSheet) {
      throw new Error("ไม่พบชีตฐานข้อมูล");
    }

    // 1. ดึงข้อมูลการตั้งค่า (แถวที่ 2 หรือ Index 1)
    var settingsData = settingSheet.getDataRange().getValues();
    var settings = {
      baseFee: "", carRate: 0, carDays: "", boatRate: 0, boatDays: "",
      bankName: "", bankAcc: "", bankAccName: ""
    };

    if (settingsData.length > 1) {
      settings = {
        baseFee: settingsData[1][0] || "",       // คอลัมน์ A: ค่าบริการส่งจีน-ไทย
        carRate: parseFloat(settingsData[1][1]) || 0, // คอลัมน์ B: นน/kg ทางรถ
        carDays: settingsData[1][2] || "",       // คอลัมน์ C: จำนวนวันทางรถ
        boatRate: parseFloat(settingsData[1][3]) || 0,// คอลัมน์ D: นน/kg ทางเรือ
        boatDays: settingsData[1][4] || "",      // คอลัมน์ E: จำนวนวันทางเรือ
        bankName: settingsData[1][5] || "",      // คอลัมน์ F: ธนาคาร
        bankAcc: settingsData[1][6] || "",       // คอลัมน์ G: เลขบัญชี
        bankAccName: settingsData[1][7] || ""    // คอลัมน์ H: ชื่อเลขบัญชี
      };
    }

    // 2. ดึงข้อมูลพัสดุของลูกค้าคนนี้
    var transportData = transportSheet.getDataRange().getValues();
    var parcels = [];
    var targetMemberId = memberId.toString().trim().toLowerCase();

    for (var i = 1; i < transportData.length; i++) {
      var rowMemberId = transportData[i][1] ? transportData[i][1].toString().trim().toLowerCase() : "";
      var status = transportData[i][11] ? transportData[i][11].toString().trim() : ""; // คอลัมน์ L: สถานะสินค้า
      
      // ดึงเฉพาะของ member นี้ (คุณสามารถเพิ่ม && status === "ถึงโกดังไทย" หากต้องการกรองเฉพาะที่ถึงไทยแล้ว)
      if (rowMemberId === targetMemberId && status === "ถึงโกดังไทย") { 
        var weightStr = transportData[i][5] ? transportData[i][5].toString().trim() : "0";
        var weight = parseFloat(weightStr) || 0;

        parcels.push({
          orderNum: transportData[i][0] ? transportData[i][0].toString() : "", 
          tracking: transportData[i][2] ? transportData[i][2].toString() : "", // คอลัมน์ C: Tracking China
          weight: weight // คอลัมน์ F: น้ำหนัก
        });
      }
    }

    return {
      success: true,
      settings: settings,
      parcels: parcels
    };

  } catch (error) {
    return { success: false, message: error.message };
  }
}