function doGet(e) {
  var page = 'index'; 
  if (e && e.parameter && e.parameter.page) {
    page = e.parameter.page;
  }
  
  try {
    return HtmlService.createTemplateFromFile(page)
        .evaluate()
        .setTitle('SUCHAR CARGO')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (error) {
    // 🛠 แก้ไขตรงนี้: โชว์ Error ออกมาตรงๆ บนหน้าเว็บแทนการแอบโหลดหน้า index
    return ContentService.createTextOutput(
      "❌ เกิดข้อผิดพลาดในการโหลดหน้า (" + page + ".html)\n\n" + 
      "รายละเอียด Error: " + error.message + 
      "\n\n(คำแนะนำ: โปรดตรวจสอบไฟล์ " + page + ".html ว่าพิมพ์แท็ก <? ?> ผิด หรือลืม include ไฟล์บางตัวหรือไม่)"
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