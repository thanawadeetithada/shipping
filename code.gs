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
    return HtmlService.createTemplateFromFile('index')
        .evaluate()
        .setTitle('SUCHAR CARGO')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}