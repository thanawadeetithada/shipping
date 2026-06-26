function getTrackingInfoData(memberId, orderId) {
  try {
    if (!memberId || !orderId) throw new Error("ข้อมูลไม่ครบถ้วน");

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var contactSheet = ss.getSheetByName("ติดต่อเรา");
    var contactData = contactSheet.getDataRange().getValues();
    var contact = { phone: "-", line: "-", lineLink: "#" };
    if (contactData.length > 1) {
      contact.phone = contactData[1][0] || "-";    
      contact.line = contactData[1][1] || "-";    
      contact.lineLink = contactData[1][2] || "#"; 
    }

    var memberSheet = ss.getSheetByName("สมาชิก");
    var memberData = memberSheet.getDataRange().getValues();
    var member = { name: "ไม่ระบุ", phone: "-", address: "ไม่มีข้อมูลที่อยู่" };
    for (var i = 1; i < memberData.length; i++) {
      if (memberData[i][0].toString() === memberId.toString()) {
        member.name = memberData[i][1] || "-";
        member.phone = memberData[i][2] || "-";
        member.address = memberData[i][3] || "-";
        break;
      }
    }

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

    var billId = orderRow[17] ? orderRow[17].toString().trim() : "";
    var totalItems = 0;
    var totalWeight = 0;
    var totalCbm = 0;
    var totalPrice = 0;

    if (billId !== "") {
      for (var j = 1; j < importData.length; j++) {
        if (importData[j][17] && importData[j][17].toString().trim() === billId) {
          totalItems++; 
          totalWeight += parseFloat(importData[j][5]) || 0; 
          totalCbm += parseFloat(importData[j][6]) || 0;
        }
      }
    } else {
      totalItems = 1;
      totalWeight = parseFloat(orderRow[5]) || 0;
      totalCbm = parseFloat(orderRow[6]) || 0;
      totalPrice = (parseFloat(orderRow[8]) || 0) + (parseFloat(orderRow[9]) || 0);
    }

    var bill = { netTotal: totalPrice.toFixed(2), status: "ยังไม่ชำระเงิน", date: "-" };
    if (billId !== "") {
      var billSheet = ss.getSheetByName("บิลชำระเงิน");
      if (billSheet) {
        var billData = billSheet.getDataRange().getValues();
        for (var k = 1; k < billData.length; k++) {
          if (billData[k][0].toString() === billId) {
            bill.netTotal = parseFloat(billData[k][6] || 0).toFixed(2);
            bill.status = billData[k][8] ? billData[k][8].toString().trim() : "รอตรวจสอบยอด";
            var bDate = billData[k][10];
            
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