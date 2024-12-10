function volunteerTimeReport() {
  // Fetch the latest form responses data
  const formResponsesData = getFormResponsesData();

  // Filter rows where column B equals the specific volunteer role
  const volunteerRows = formResponsesData.filter((row, index) => {
    return index > 0 && row[1] === "Northside resource fair volunteer (NHS, NJHS, etc.)";
  });

  // Group rows by the values in column H
  const groupedData = {};
  volunteerRows.forEach(row => {
    const key = String(row[7]).toLowerCase(); // Column H
    if (!groupedData[key]) {
      groupedData[key] = [];
    }
    groupedData[key].push(row);
  });

  // Process each group to find "IN" and "OUT" rows and calculate time worked
  const resultData = [["E or S#", "Name", "Campus/Department", "Time In", "Time Out", "Total Time Volunteered (hours.minutes)"]];
  for (const [key, rows] of Object.entries(groupedData)) {
    const inRow = rows.find(row => row[10] === "IN (You just got to the fair and are ready to begin your shift.)"); // Column K
    const outRow = rows.find(row => row[10] === "OUT (You finished your shift and are going home.)"); // Column K

    if (inRow && outRow) {
      // Both IN and OUT rows exist
      const inTimestamp = inRow[0]; // Column A
      const outTimestamp = outRow[0]; // Column A
      if (inTimestamp instanceof Date && outTimestamp instanceof Date) {
        const formattedInTimestamp = Utilities.formatDate(inTimestamp, Session.getScriptTimeZone(), "hh:mm:ss a");
        const formattedOutTimestamp = Utilities.formatDate(outTimestamp, Session.getScriptTimeZone(), "hh:mm:ss a");
        const name = inRow[8]; // Column I (Name)
        const campus = inRow[9]; // Column J (Campus/Department)
        const timeWorked = (outTimestamp - inTimestamp) / (1000 * 60 * 60); // Convert ms to hours
        resultData.push([key, name, campus, formattedInTimestamp, formattedOutTimestamp, timeWorked.toFixed(2)]);
      }
    } else if (inRow) {
      // Missing OUT row
      const inTimestamp = inRow[0]; // Column A
      const formattedInTimestamp = Utilities.formatDate(inTimestamp, Session.getScriptTimeZone(), "hh:mm:ss a");
      const name = inRow[8]; // Column I (Name)
      const campus = inRow[9]; // Column J (Campus/Department)
      resultData.push([key, name, campus, formattedInTimestamp, "Did not clock out", "Unable to calculate"]);
    } else if (outRow) {
      // Missing IN row
      const outTimestamp = outRow[0]; // Column A
      const formattedOutTimestamp = Utilities.formatDate(outTimestamp, Session.getScriptTimeZone(), "hh:mm:ss a");
      const name = outRow[8]; // Column I (Name)
      const campus = outRow[9]; // Column J (Campus/Department)
      resultData.push([key, name, campus, "Did not clock in", formattedOutTimestamp, "Unable to calculate"]);
    }
  }

  // Output the calculated data to a new sheet
  const outputSheetName = "Volunteer Time Report";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let outputSheet = ss.getSheetByName(outputSheetName);

  if (!outputSheet) {
    outputSheet = ss.insertSheet(outputSheetName);
  } else {
    outputSheet.clear();
  }

  // Write the result data to the new sheet
  outputSheet.getRange(1, 1, resultData.length, resultData[0].length).setValues(resultData);
}
