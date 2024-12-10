function getFormResponsesData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formResponsesSheet = ss.getSheetByName('Form Responses 1');
  return formResponsesSheet.getRange(1, 1, formResponsesSheet.getLastRow(), 22).getValues();
}