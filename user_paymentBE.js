var SETTING_SHEET_NAME_PAYMENT = "ตั้งค่า";
var BILL_SHEET_NAME_PAYMENT = "บิลชำระเงิน";
var IMPORT_SHEET_NAME_PAYMENT = "ข้อมูลขนส่ง";
var RATE_SHEET_NAME_PAYMENT = "อัตราค่าขนส่ง"; 

function extractNumberForCalculate(str) {
  if (!str) return 0;
  var match = str.toString().match(/\d+(\.\d+)?/);
  return match ? parseFloat(match[0]) : 0;
}

function getPaymentPageData(memberId) {
  try {
    if (!memberId) throw new Error("ไม่พบข้อมูลสมาชิก");

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var settingSheet = ss.getSheetByName(SETTING_SHEET_NAME_PAYMENT);
    var transportSheet = ss.getSheetByName(IMPORT_SHEET_NAME_PAYMENT);
    var rateSheet = ss.getSheetByName(RATE_SHEET_NAME_PAYMENT);

    if (!settingSheet || !transportSheet || !rateSheet) {
      throw new Error("ไม่พบชีตฐานข้อมูล");
    }

    var settingsData = settingSheet.getDataRange().getValues();
    var settings = { baseFee: "", bankName: "", bankAcc: "", bankAccName: "", qrCodeUrl: "" };

    if (settingsData.length > 1) {
      var rawQrUrl = settingsData[1][8] ? settingsData[1][8].toString().trim() : "";
      var formattedQrUrl = rawQrUrl;
      
      if (rawQrUrl.indexOf("drive.google.com") !== -1) {
        var match = rawQrUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
        if (!match) match = rawQrUrl.match(/id=([a-zA-Z0-9_-]+)/);
        
        if (match && match[1]) {
          formattedQrUrl = "https://lh3.googleusercontent.com/d/" + match[1];
        }
      }

      settings = {
        baseFee: settingsData[1][0] || "",             
        bankName: settingsData[1][5] || "",            
        bankAcc: settingsData[1][6] || "",             
        bankAccName: settingsData[1][7] || "",
        qrCodeUrl: formattedQrUrl
      };
    }

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

    var transportData = transportSheet.getDataRange().getValues();
    var parcels = [];
    var targetMemberId = memberId.toString().trim().toLowerCase();

    for (var i = 1; i < transportData.length; i++) {
      var rowMemberId = transportData[i][1] ? transportData[i][1].toString().trim().toLowerCase() : "";
      var rowBillId = transportData[i][17] ? transportData[i][17].toString().trim() : "";
      
      if (rowMemberId === targetMemberId && rowBillId === "") { 
        var weightStr = transportData[i][6] ? transportData[i][6].toString().trim() : "0";
        var cbmStr = transportData[i][7] ? transportData[i][7].toString().trim() : "0";
        
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

function submitPaymentData(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var billSheet = ss.getSheetByName(BILL_SHEET_NAME_PAYMENT);
    var importSheet = ss.getSheetByName(IMPORT_SHEET_NAME_PAYMENT);

    if (!billSheet || !importSheet) throw new Error("ไม่พบชีตสำหรับบันทึกข้อมูล");

    var slipUrl = "";
    if (payload.slipBase64) {
      var blob = Utilities.newBlob(Utilities.base64Decode(payload.slipBase64), payload.slipMime, payload.slipName);
      var folder = DriveApp.getFolderById("1bJHkMVxMdXPDJ3XRZJzcTRfrNYYXLbLm");
      var file = folder.createFile(blob);
      try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch(e) {}
      slipUrl = file.getUrl();
    }

    var timestampStr = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyyMMddHHmmss");
    var billId = "INV" + timestampStr;
    var currentTime = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss");

    billSheet.appendRow([
      billId,                          
      payload.memberId,            
      payload.orderList.join(", "),     
      payload.totalTransportFee,     
      "",                              
      "",                            
      payload.finalTotal,            
      slipUrl,                        
      "รอตรวจสอบยอด",               
      "",                             
      currentTime                    
    ]);


    var importData = importSheet.getDataRange().getValues();
    var targetOrders = payload.orderList;

    for (var i = 1; i < importData.length; i++) {
      var rowOrderNum = importData[i][0] ? importData[i][0].toString().trim() : "";
      
      if (targetOrders.indexOf(rowOrderNum) !== -1) {
        importSheet.getRange(i + 1, 5).setValue(payload.shippingType);  
        importSheet.getRange(i + 1, 6).setValue(payload.productType); 
        importSheet.getRange(i + 1, 18).setValue(billId);             
      }
    }

    return { success: true, message: "บันทึกข้อมูลและอัปโหลดสลิปเรียบร้อยแล้ว!" };

  } catch (error) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + error.message };
  }
}