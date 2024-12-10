function rankCampuses() {
  // Fetch the latest form responses data
  const formResponsesData = getFormResponsesData();

  // Filter rows where column C is not empty
  const campuses = formResponsesData
    .slice(1) // Skip the header row
    .filter(row => row[2]) // Only include rows where column C has a value
    .map(row => row[2]); // Extract column C (campus)

  // Count occurrences of each campus
  const campusCounts = campuses.reduce((counts, campus) => {
    counts[campus] = (counts[campus] || 0) + 1;
    return counts;
  }, {});

  // Convert the counts object to an array of [campus, count] pairs
  const campusArray = Object.entries(campusCounts);

  // Sort by count in descending order, breaking ties alphabetically
  campusArray.sort((a, b) => {
    if (b[1] === a[1]) {
      return a[0].localeCompare(b[0]); // Alphabetical order for ties
    }
    return b[1] - a[1]; // Descending order by count
  });

  // Prepare the result data
  const resultData = [["Campus", "Count"]];
  campusArray.forEach(([campus, count]) => {
    resultData.push([campus, count]);
  });

  // Output the ranked campuses to a new sheet
  const outputSheetName = "Campus Numbers";
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
