function doGet(e) {
  var page = e && e.parameter && e.parameter.page ? e.parameter.page : 'index'; 
  console.log("กำลังโหลดไฟล์: " + page + ".html");
  
  try {
    // 1. ใช้ Template เสมอ เพื่อให้ใช้งาน <?!= include(...) ?> ได้
    return HtmlService.createTemplateFromFile(page)
        .evaluate()
        .setTitle('SUCHAR CARGO')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (error) {
    return ContentService.createTextOutput(
      "❌ ไม่พบหน้า: " + page + ".html\n" + 
      "Error: " + error.message
    );
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

// เปลี่ยนจากของเดิม ให้เป็นแบบนี้
// function getScriptUrl() {
//   // เปลี่ยนลิงก์ด้านล่างให้เป็น URL GitHub Pages ของคุณจริงๆ
//   // ตัวอย่างเช่น: "https://yourusername.github.io/sucharcargo/"
//   return "https://sucharforwork.github.io/shipping/"; 
// }