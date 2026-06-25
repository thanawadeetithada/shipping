/**
 * ดึงข้อมูลพัสดุสำหรับหน้าสาธารณะ (ไม่เช็ค Session/Auth ใดๆ ทั้งสิ้น)
 * ค้นหาได้ทั้ง เลขออเดอร์ หรือ Tracking จีน
 */
function getTrackingInfoData_Public(orderId) {
  try {
    // 1. เช็คแค่ว่ามีการพิมพ์เลขค้นหามาหรือไม่
    if (!orderId || String(orderId).trim() === "") {
      throw new Error("กรุณาระบุเลขออเดอร์ หรือ Tracking Number");
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    // ทำความสะอาดและแปลงตัวอักษรเป็นพิมพ์ใหญ่เพื่อป้องกันปัญหาพิมพ์ผิด (Case Insensitive)
    var orderQuery = String(orderId).trim().toUpperCase();

    // 2. ดึงข้อมูลการติดต่อ (ชีต "ติดต่อเรา")
    var contactSheet = ss.getSheetByName("ติดต่อเรา");
    var contactData = contactSheet.getDataRange().getValues();
    var contact = { phone: "-", line: "-", lineLink: "#" };
    if (contactData.length > 1) {
      contact.phone = contactData[1][0] || "-";    
      contact.line = contactData[1][1] || "-";     
      contact.lineLink = contactData[1][2] || "#"; 
    }

    // 3. ค้นหาข้อมูลพัสดุ (ชีต "ข้อมูลขนส่ง") 
    var importSheet = ss.getSheetByName("ข้อมูลขนส่ง");
    var importData = importSheet.getDataRange().getValues();
    var orderRow = null;
    var actualMemberId = ""; 
    
    for (var i = 1; i < importData.length; i++) {
      // ค้นหาจากเลขออเดอร์ (คอลัมน์ A) หรือ Tracking จีน (คอลัมน์ C)
      var cellOrder = String(importData[i][0] || "").trim().toUpperCase();
      var cellTracking = String(importData[i][2] || "").trim().toUpperCase();

      if (cellOrder === orderQuery || cellTracking === orderQuery) {
        orderRow = importData[i];
        actualMemberId = importData[i][1] ? String(importData[i][1]).trim() : ""; 
        break;
      }
    }

    if (!orderRow) {
      throw new Error("ไม่พบข้อมูลพัสดุนี้ในระบบ กรุณาตรวจสอบรหัสอีกครั้ง");
    }

    // 4. ดึงข้อมูลสมาชิก (ชีต "สมาชิก") เพื่อเอาที่อยู่จัดส่ง
    var memberSheet = ss.getSheetByName("สมาชิก");
    var memberData = memberSheet.getDataRange().getValues();
    var member = { name: "ไม่ระบุ", phone: "-", address: "ไม่มีข้อมูลที่อยู่" };
    
    if (actualMemberId !== "") {
      for (var m = 1; m < memberData.length; m++) {
        if (String(memberData[m][0]).trim() === actualMemberId) {
          member.name = memberData[m][1] || "-";
          member.phone = memberData[m][2] || "-";
          member.address = memberData[m][3] || "-"; 
          break;
        }
      }
    }

    // 5. คำนวณจำนวนชิ้น/น้ำหนัก/ปริมาตร รวมใน 1 บิล
    var billId = orderRow[17] ? String(orderRow[17]).trim() : ""; 
    var totalItems = 0;
    var totalWeight = 0;
    var totalCbm = 0;
    var totalPrice = 0;

    if (billId !== "") {
      for (var j = 1; j < importData.length; j++) {
        if (importData[j][17] && String(importData[j][17]).trim() === billId) {
          totalItems++; 
          totalWeight += parseFloat(importData[j][6]) || 0; // น้ำหนักอยู่คอลัมน์ G (Index 6)
          totalCbm += parseFloat(importData[j][7]) || 0;    // CBM อยู่คอลัมน์ H (Index 7)
        }
      }
    } else {
      totalItems = 1;
      totalWeight = parseFloat(orderRow[6]) || 0; 
      totalCbm = parseFloat(orderRow[7]) || 0;
      totalPrice = (parseFloat(orderRow[8]) || 0) + (parseFloat(orderRow[9]) || 0); 
    }

    // 6. ดึงข้อมูลบิลชำระเงิน (ถ้ามีบิลแล้ว)
    var bill = { netTotal: totalPrice.toFixed(2), status: "ยังไม่ชำระเงิน", date: "-" };
    if (billId !== "") {
      var billSheet = ss.getSheetByName("บิลชำระเงิน");
      if (billSheet) {
        var billData = billSheet.getDataRange().getValues();
        for (var k = 1; k < billData.length; k++) {
          if (String(billData[k][0]).trim() === billId) {
            bill.netTotal = parseFloat(billData[k][6] || 0).toFixed(2); 
            bill.status = billData[k][8] ? String(billData[k][8]).trim() : "รอตรวจสอบยอด"; 
            var bDate = billData[k][10]; 
            
            if(bDate) {
              var bd = new Date(bDate);
              if(!isNaN(bd.getTime())) {
                var mths = ["ม.ค.","ก.พ.","มี.ค.","เม.ย.","พ.ค.","มิ.ย.","ก.ค.","ส.ค.","ก.ย.","ต.ค.","พ.ย.","ธ.ค."];
                bill.date = bd.getDate() + " " + mths[bd.getMonth()] + " " + (bd.getFullYear() + 543) + " " + 
                            ("0" + bd.getHours()).slice(-2) + ":" + ("0" + bd.getMinutes()).slice(-2) + " น.";
              }
            }
            break;
          }
        }
      }
    }

    // 7. แปลงวันที่สำหรับ Timeline
    var formatTimelineDate = function(d) {
      if (!d) return null;
      var date = new Date(d);
      if (isNaN(date.getTime())) return null;
      var months = ["ม.ค.","ก.พ.","มี.ค.","เม.ย.","พ.ค.","มิ.ย.","ก.ค.","ส.ค.","ก.ย.","ต.ค.","พ.ย.","ธ.ค."];
      return { 
        dateStr: date.getDate() + " " + months[date.getMonth()] + " " + (date.getFullYear() + 543), 
        timeStr: ("0" + date.getHours()).slice(-2) + ":" + ("0" + date.getMinutes()).slice(-2)
      };
    };

    var latestStatus = orderRow[11] ? String(orderRow[11]).trim() : "แจ้งนำเข้าแล้ว";

    // 8. ส่งออกข้อมูล
    return {
      success: true,
      contact: contact,
      member: member,
      tracking: {
        orderNum: String(orderRow[0]).trim(),
        trackingId: String(orderRow[2]).trim(),
        statusText: latestStatus,
        timeline: {
          step1: formatTimelineDate(orderRow[12]), 
          step2: formatTimelineDate(orderRow[13]), 
          step3: formatTimelineDate(orderRow[14]), 
          step4: formatTimelineDate(orderRow[15]), 
          step5: formatTimelineDate(orderRow[16])  
        }
      },
      summary: {
        items: totalItems,
        weight: totalWeight.toFixed(2),
        cbm: totalCbm.toFixed(2),
        netTotal: bill.netTotal,
        paymentStatus: bill.status,
        paymentDate: bill.date
      }
    };

  } catch (error) {
    return { success: false, message: error.message };
  }
}