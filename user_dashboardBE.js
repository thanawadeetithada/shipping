/**
 * ดึงข้อมูลสรุปสำหรับหน้า Dashboard ของลูกค้า
 */
function getDashboardData(memberId) {
  try {
    if (!memberId) throw new Error("ไม่พบข้อมูลผู้ใช้งาน กรุณาเข้าสู่ระบบใหม่");

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var transportSheet = ss.getSheetByName("ข้อมูลขนส่ง");
    var billSheet = ss.getSheetByName("บิลชำระเงิน");

    var transportData = transportSheet.getDataRange().getValues();
    var billData = billSheet ? billSheet.getDataRange().getValues() : [];

    // ดึงข้อมูลบิลมาทำเป็น Dictionary เพื่อง่ายต่อการค้นหา
    var billDict = {};
    for (var i = 1; i < billData.length; i++) {
      var bId = billData[i][0] ? billData[i][0].toString().trim() : "";
      if (bId) {
        billDict[bId] = {
          total: parseFloat(billData[i][6] || 0), // ยอดสุทธิ
          status: billData[i][8] ? billData[i][8].toString().trim() : "รอตรวจสอบยอด"
        };
      }
    }

    var summary = { waitChina: 0, comingThai: 0, waitPay: 0, successMonth: 0 };
    var chinaItems = [];
    var shipments = [];

    var currentMonth = new Date().getMonth();
    var currentYear = new Date().getFullYear();

    // วนลูปเช็คข้อมูลขนส่งของ User คนนี้
    for (var j = 1; j < transportData.length; j++) {
      var row = transportData[j];
      
      // เช็ค Member_ID คอลัมน์ B (Index 1)
      if (row[1] && row[1].toString().trim() === memberId.toString().trim()) {
        var orderId = row[0] ? row[0].toString() : "";
        var chinaTrack = row[2] ? row[2].toString() : "-";
        var details = row[4] ? row[4].toString() : "สินค้าไม่ระบุชื่อ";
        var weight = parseFloat(row[5]) || 0;
        var cbm = parseFloat(row[6]) || 0;
        var totalCost = parseFloat(row[10]) || 0;
        var status = row[11] ? row[11].toString().trim() : "แจ้งนำเข้าแล้ว";
        var deliveryDate = row[16]; // คอลัมน์วันที่จัดส่งสำเร็จ (Q)
        var billId = row[17] ? row[17].toString().trim() : "";

        var displayStatus = status;
        var finalTotal = totalCost;

        // ถ้าออเดอร์นี้ถูกผูกบิลแล้ว ให้ใช้สถานะการชำระเงินจากบิลมาแสดงนำ
        if (billId && billDict[billId]) {
          var bStatus = billDict[billId].status;
          if (bStatus === "รอชำระเงิน" || bStatus === "รอตรวจสอบยอด" || bStatus === "ชำระแล้ว") {
             displayStatus = bStatus;
          }
        }

        // --- คำนวณ Summary 4 กล่องด้านบน ---
        if (status === "แจ้งนำเข้าแล้ว" || status === "รอเข้าโกดังจีน") {
          summary.waitChina++;
        } else if (status === "กำลังขนส่งมาไทย") {
          summary.comingThai++;
        } else if (status === "ถึงโกดังไทย" || displayStatus === "รอชำระเงิน" || displayStatus === "รอตรวจสอบยอด") {
          summary.waitPay++;
        } else if (status === "จัดส่งสำเร็จ") {
          // เช็คว่าเป็นของเดือนนี้หรือไม่
          if (deliveryDate) {
            var dDate = new Date(deliveryDate);
            if (dDate.getMonth() === currentMonth && dDate.getFullYear() === currentYear) {
               summary.successMonth++;
            }
          } else {
             summary.successMonth++; // ถ้าไม่มีวันที่ ถือว่านับรวมไปก่อน
          }
        }

        // --- แยกหมวดหมู่ลงรายการ ---
        if (status === "ถึงโกดังจีน") {
           // 1. สินค้าในโกดังจีน รอแพ็ครวมบิล
           chinaItems.push({
             id: orderId,
             name: details,
             track: chinaTrack,
             weight: weight.toFixed(2),
             cbm: cbm.toFixed(2)
           });
        } 
        else if (status !== "แจ้งนำเข้าแล้ว" && status !== "รอเข้าโกดังจีน" && status !== "จัดส่งสำเร็จ") {
           // 2. สินค้าที่กำลังเดินทาง, ถึงไทยแล้ว, หรือรอชำระเงิน (Active Shipments)
           var totalStr = finalTotal > 0 ? finalTotal.toFixed(2) + " ฿" : "รอกำหนดราคา";
           if (billId && billDict[billId]) {
              totalStr = "รวมในบิล " + billId + " (" + billDict[billId].total.toFixed(2) + " ฿)";
           }

           shipments.push({
             id: orderId, // ใช้ Order_ID แสดงเป็นรหัสติดตาม
             status: displayStatus,
             weight: weight.toFixed(2),
             totalDisplay: totalStr
           });
        }
      }
    }

    // Sort active shipments ให้รายการที่ต้องดำเนินการ (รอชำระ, ถึงไทย) ขึ้นมาก่อน
    var statusPriority = {
      "รอชำระเงิน": 1,
      "ถึงโกดังไทย": 2,
      "รอตรวจสอบยอด": 3,
      "กำลังขนส่งมาไทย": 4,
      "ชำระแล้ว": 5
    };

    shipments.sort(function(a, b) {
       var p1 = statusPriority[a.status] || 99;
       var p2 = statusPriority[b.status] || 99;
       return p1 - p2;
    });

    return {
      success: true,
      summary: summary,
      chinaItems: chinaItems,
      shipments: shipments
    };

  } catch (error) {
    return { success: false, message: error.message };
  }
}