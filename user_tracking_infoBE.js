/**
 * ดึงข้อมูลพัสดุ 1 รายการ และข้อมูลบิล/สมาชิก/ติดต่อ เพื่อแสดงในหน้าสถานะแบบละเอียด
 */
function getTrackingInfoData(memberId, orderId) {
  try {
    if (!memberId || !orderId) throw new Error("ข้อมูลไม่ครบถ้วน");

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. ดึงข้อมูลการติดต่อ (ชีต "ติดต่อเรา")
    var contactSheet = ss.getSheetByName("ติดต่อเรา");
    var contactData = contactSheet.getDataRange().getValues();
    var contact = { phone: "-", line: "-", lineLink: "#" };
    if (contactData.length > 1) {
      contact.phone = contactData[1][0] || "-";    // คอลัมน์ A
      contact.line = contactData[1][1] || "-";     // คอลัมน์ B
      contact.lineLink = contactData[1][2] || "#"; // คอลัมน์ C
    }

    // 2. ดึงข้อมูลสมาชิก (ชีต "สมาชิก") เพื่อเอาที่อยู่จัดส่ง
    var memberSheet = ss.getSheetByName("สมาชิก");
    var memberData = memberSheet.getDataRange().getValues();
    var member = { name: "ไม่ระบุ", phone: "-", address: "ไม่มีข้อมูลที่อยู่" };
    for (var i = 1; i < memberData.length; i++) {
      if (memberData[i][0].toString() === memberId.toString()) {
        member.name = memberData[i][1] || "-";
        member.phone = memberData[i][2] || "-";
        member.address = memberData[i][3] || "-"; // คอลัมน์ D: ที่อยู่
        break;
      }
    }

    // 3. ค้นหาข้อมูลพัสดุ (ชีต "ข้อมูลขนส่ง")
    var importSheet = ss.getSheetByName("ข้อมูลขนส่ง");
    var importData = importSheet.getDataRange().getValues();
    var orderRow = null;
    
    for (var i = 1; i < importData.length; i++) {
      if (importData[i][0].toString() === orderId.toString() && importData[i][1].toString() === memberId.toString()) {
        orderRow = importData[i];
        break;
      }
    }

    if (!orderRow) {
      throw new Error("ไม่พบข้อมูลพัสดุนี้ หรือไม่ใช่พัสดุของคุณ");
    }

    // อ่าน Bill_ID จากคอลัมน์ R (Index 17)
    var billId = orderRow[17] ? orderRow[17].toString().trim() : "";
    var totalItems = 0;
    var totalWeight = 0;
    var totalCbm = 0;
    var totalPrice = 0;

    // คำนวณจำนวนชิ้น/น้ำหนัก/ปริมาตร รวมใน 1 บิล
    if (billId !== "") {
      for (var j = 1; j < importData.length; j++) {
        if (importData[j][17] && importData[j][17].toString().trim() === billId) {
          totalItems++; // นับจำนวน Order ในบิลนี้
          totalWeight += parseFloat(importData[j][5]) || 0; // รวมน้ำหนัก
          totalCbm += parseFloat(importData[j][6]) || 0;    // รวม CBM
        }
      }
    } else {
      totalItems = 1;
      totalWeight = parseFloat(orderRow[5]) || 0;
      totalCbm = parseFloat(orderRow[6]) || 0;
      totalPrice = (parseFloat(orderRow[8]) || 0) + (parseFloat(orderRow[9]) || 0); // ราคาเดี่ยวกรณียังไม่มีบิล
    }

    // 4. ดึงข้อมูลบิลชำระเงิน (ถ้ามีบิลแล้ว)
    var bill = { netTotal: totalPrice.toFixed(2), status: "ยังไม่ชำระเงิน", date: "-" };
    if (billId !== "") {
      var billSheet = ss.getSheetByName("บิลชำระเงิน");
      if (billSheet) {
        var billData = billSheet.getDataRange().getValues();
        for (var k = 1; k < billData.length; k++) {
          if (billData[k][0].toString() === billId) {
            bill.netTotal = parseFloat(billData[k][6] || 0).toFixed(2); // คอลัมน์ G: ยอดสุทธิ
            bill.status = billData[k][8] ? billData[k][8].toString().trim() : "รอตรวจสอบยอด"; // คอลัมน์ I: สถานะบิล
            var bDate = billData[k][10]; // คอลัมน์ K: วันที่ชำระเงิน
            
            // แปลงรูปแบบวันที่ชำระเงิน
            if(bDate) {
              var bd = new Date(bDate);
              var mths = ["ม.ค.","ก.พ.","มี.ค.","เม.ย.","พ.ค.","มิ.ย.","ก.ค.","ส.ค.","ก.ย.","ต.ค.","พ.ย.","ธ.ค."];
              bill.date = bd.getDate() + " " + mths[bd.getMonth()] + " " + (bd.getFullYear() + 543) + " " + 
                          ("0" + bd.getHours()).slice(-2) + ":" + ("0" + bd.getMinutes()).slice(-2) + " น.";
            }
            break;
          }
        }
      }
    }

    // ฟังก์ชันช่วยแปลงวันที่สำหรับ Timeline (M - Q)
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

    var latestStatus = orderRow[11] ? orderRow[11].toString() : "แจ้งนำเข้าแล้ว";

    return {
      success: true,
      contact: contact,
      member: member,
      tracking: {
        orderNum: orderRow[0].toString(),
        trackingId: orderRow[2].toString(),
        statusText: latestStatus,
        timeline: {
          step1: formatTimelineDate(orderRow[12]), // M: วันที่แจ้งนำเข้า
          step2: formatTimelineDate(orderRow[13]), // N: วันที่ถึงโกดังจีน
          step3: formatTimelineDate(orderRow[14]), // O: วันที่กำลังขนส่งมาไทย
          step4: formatTimelineDate(orderRow[15]), // P: วันที่ถึงโกดังไทย
          step5: formatTimelineDate(orderRow[16])  // Q: วันที่จัดส่งสำเร็จ
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