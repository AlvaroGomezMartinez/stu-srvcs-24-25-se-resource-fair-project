function filterParentResponses() {
  // Fetch the latest Form Responses 1 data
  const formResponsesData = getFormResponsesData();
  
  // Extract headers from the data
  const headers = formResponsesData[0];
  
  // Filter the rows where column B equals "parent/guardian of a Northside student(s)"
  const filteredData = formResponsesData.filter((row, index) => {
    return index === 0 || row[1] === "parent/guardian of a Northside student(s)";
  });

  // Limit data to columns A:G
  const limitedData = filteredData.map(row => row.slice(0, 7));

  // Replace "Check one or both" with "Yes" in columns E (index 4) and F (index 5)
  const updatedData = limitedData.map((row, index) => {
    if (index > 0) { // Skip the header row
      if (row[4] === "Check one or both") row[4] = "Yes"; // Update column E
      if (row[5] === "Check one or both") row[5] = "Yes"; // Update column F
      
      // Format the timestamp in column A (index 0)
      if (row[0] instanceof Date) {
        row[0] = Utilities.formatDate(row[0], Session.getScriptTimeZone(), "MM/dd/yy hh:mm a");
      }
    }
    return row;
  });

  // Output the filtered data to a new sheet
  const outputSheetName = "Parents";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let outputSheet = ss.getSheetByName(outputSheetName);
  
  if (!outputSheet) {
    outputSheet = ss.insertSheet(outputSheetName);
  } else {
    outputSheet.clear();
  }
  
  // Write the filtered and updated data to the new sheet
  outputSheet.getRange(1, 1, updatedData.length, updatedData[0].length).setValues(updatedData);

  // Apply text wrapping to the headers in columns E and F
  const headerRange = outputSheet.getRange(1, 5, 1, 2); // E1:F1
  headerRange.setWrap(true);
}
