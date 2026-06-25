function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    var action = e.parameter.action;
    
    if (action === "getTracking") {
      var memberId = e.parameter.memberId || "";
      var orderId = e.parameter.orderId || "";  
      var result = getTrackingInfoData(memberId, orderId);
      
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }
        
    return ContentService.createTextOutput(JSON.stringify({success: false, message: "Invalid Action"}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var page = e && e.parameter && e.parameter.page ? e.parameter.page : 'index'; 
  console.log("กำลังโหลดไฟล์: " + page + ".html");
  
  try {
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
  return "https://sucharforwork.github.io/shipping/"; 
}

// function getScriptUrl() { //หน้า appscript
//   return ScriptApp.getService().getUrl();
// }
