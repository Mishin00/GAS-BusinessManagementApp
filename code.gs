const SPREADSHEET_ID = '1QDcfmpXfjtaxgptuL9LwctPk7fUgVtc_Vz_ySyf027k'
const SHEET_BUSINESSES = 'Businesses';
const SHEET_TASKS = 'Tasks';

function doGet(e) {
  let page = e.parameter.page;
  if (!page){page = 'index'}
  return HtmlService.createTemplateFromFile(page).evaluate();
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSheet(name) {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name);
}

function getDataRangeValues(name) {
  const sheet = getSheet(name);
  return sheet.getDataRange().getValues();
}

function getHeaderIndexMap(headers) {
  return headers.reduce((map, h, i) => {
    map[h] = i;
    return map;
  }, {});
}

function getAppUrl() {
  return ScriptApp.getService().getUrl();
}
