function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Manifest Tools')
    .addItem('New Manifest', 'showSheetSelector')
    .addItem('Sum Selected Weights', 'sumSelectedWeights')
    .addToUi();
}

// Show the custom dialog for sheet selection
function showSheetSelector() {
  var html = HtmlService.createHtmlOutputFromFile('SheetSelector')
    .setWidth(400)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Select Source and Destination Sheets');
}

// Return all sheet names to the HTML dialog
function getSheetNames() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets().map(sheet => sheet.getName());
}

// Process the sheet names selected by the user in the dialog
function processSheetSelection(sourceSheetNames, masterSheetName) {
  if (!masterSheetName || sourceSheetNames.length === 0) {
    SpreadsheetApp.getUi().alert('Please select at least one source sheet and a destination sheet.');
    return;
  }
  // Call your manifest creation function with dynamic sheet names
  CreateNewTourManifest(sourceSheetNames, masterSheetName);
}

// Modified CreateNewTourManifest to accept dynamic sheet names
function CreateNewTourManifest(sourceSheetNames, masterSheetName) {
  // Data starts from row 3 in source sheets
  const startRow = 3;
  const endRow = 1000; // Last row to update

  // Column indices for A–J (1–10) and P–S (16–19)
  const sourceAtoJ = [0,1,2,3,4,5,6,7,8,9];     // A–J (0-based)
  const sourcePtoS = [15,16,17,18];             // P–S (0-based)
  const targetAtoJ = [1,2,3,4,5,6,7,8,9,10];    // A–J (1-based)
  const targetPtoS = [16,17,18,19];             // P–S (1-based)

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let combinedAtoJ = [];
  let combinedAtoJColors = [];
  let combinedPtoS = [];
  let combinedColBBold = [];
  let combinedColBItalic = [];

  // Gather data and colors from each source sheet
  sourceSheetNames.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow < startRow) return; // Skip if no data

    const numCols = 19; // A–S
    const numRows = lastRow - startRow + 1;
    const dataRange = sheet.getRange(startRow, 1, numRows, numCols);
    const values = dataRange.getValues();
    const colors = dataRange.getBackgrounds();

    // For bold/italic in column B (col 2)
    const colBRange = sheet.getRange(startRow, 2, numRows, 1);
    const bolds = colBRange.getFontWeights();
    const italics = colBRange.getFontStyles();

    // Filter out rows where Column B is blank
    for (let i = 0; i < values.length; i++) {
      if (values[i][1] !== "" && values[i][1] !== null) { // Column B is index 1
        // Extract A–J values and colors
        let rowAtoJ = [];
        let rowAtoJColor = [];
        sourceAtoJ.forEach(idx => {
          rowAtoJ.push(values[i][idx]);
          rowAtoJColor.push(colors[i][idx]);
        });
        combinedAtoJ.push(rowAtoJ);
        combinedAtoJColors.push(rowAtoJColor);

        // Extract P–S values only
        let rowPtoS = [];
        sourcePtoS.forEach(idx => {
          rowPtoS.push(values[i][idx]);
        });
        combinedPtoS.push(rowPtoS);

        // Extract bold/italic for column B
        combinedColBBold.push([bolds[i][0]]);
        combinedColBItalic.push([italics[i][0]]);
      }
    }
  });

  // Limit how many rows we update so we don't touch 1001 or 1002
  const maxRowsToWrite = Math.min(combinedAtoJ.length, endRow - startRow + 1);

  // Write to master sheet
  let masterSheet = ss.getSheetByName(masterSheetName);
  if (!masterSheet) {
    masterSheet = ss.insertSheet(masterSheetName);
  } else {
    // Clear only values in rows 3–1000 for columns A–J and P–S (do NOT clear formatting)
    targetAtoJ.concat(targetPtoS).forEach(function(colIdx) {
      masterSheet.getRange(startRow, colIdx, endRow - startRow + 1, 1).clearContent();
    });
  }

  // Write combined data and colors starting from row 3, only to columns A–J and P–S, up to row 1000
  if (maxRowsToWrite > 0) {
    // Columns A–J: set values and background
    for (let c = 0; c < targetAtoJ.length; c++) {
      let colData = combinedAtoJ.slice(0, maxRowsToWrite).map(row => [row[c]]);
      let colColors = combinedAtoJColors.slice(0, maxRowsToWrite).map(row => [row[c]]);
      masterSheet.getRange(startRow, targetAtoJ[c], maxRowsToWrite, 1).setValues(colData);
      // Only set background for A–J
      masterSheet.getRange(startRow, targetAtoJ[c], maxRowsToWrite, 1).setBackgrounds(colColors);
    }
    // Columns P–S: set values only (do not touch background)
    for (let c = 0; c < targetPtoS.length; c++) {
      let colData = combinedPtoS.slice(0, maxRowsToWrite).map(row => [row[c]]);
      masterSheet.getRange(startRow, targetPtoS[c], maxRowsToWrite, 1).setValues(colData);
    }
    // Column B (col 2): set bold and italic
    masterSheet.getRange(startRow, 2, maxRowsToWrite, 1).setFontWeights(combinedColBBold.slice(0, maxRowsToWrite));
    masterSheet.getRange(startRow, 2, maxRowsToWrite, 1).setFontStyles(combinedColBItalic.slice(0, maxRowsToWrite));
  }
}

// === Sum Selected to Col E ===
function sumSelectedAndAddToColE() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getActiveRange();
  var values = range.getValues();
  var sum = 0;

  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[0].length; j++) {
      var val = values[i][j];
      if (typeof val === "number" && !isNaN(val)) {
        sum += val;
      }
    }
  }

  var topRow = range.getRow();
  sheet.getRange(topRow, 5).setValue(sum);
}

// === Compare Two Sheets ===
function compareColumnBWithPrompt() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Prompt for first sheet name
  var result1 = ui.prompt('Enter the name of the FIRST sheet to compare:');
  if (result1.getSelectedButton() !== ui.Button.OK) return;
  var sheetName1 = result1.getResponseText();

  // Prompt for second sheet name
  var result2 = ui.prompt('Enter the name of the SECOND sheet to compare:');
  if (result2.getSelectedButton() !== ui.Button.OK) return;
  var sheetName2 = result2.getResponseText();

  var sheet1 = ss.getSheetByName(sheetName1);
  var sheet2 = ss.getSheetByName(sheetName2);

  if (!sheet1 || !sheet2) {
    ui.alert('One or both sheet names are invalid.');
    return;
  }

  // Get data from Column B (second column)
  var colB1 = sheet1.getRange(1, 2, sheet1.getLastRow(), 1).getValues();
  var colB2 = sheet2.getRange(1, 2, sheet2.getLastRow(), 1).getValues();

  // Clear previous highlights in Column B
  // sheet1.getRange(1, 2, sheet1.getLastRow(), 1).setBackground(null);

  // Compare and highlight differences
  var numRows = Math.min(colB1.length, colB2.length);
  for (var i = 0; i < numRows; i++) {
    if (colB1[i][0] !== colB2[i][0]) {
      sheet1.getRange(i + 1, 2).setBackground('#ffeb3b'); // Yellow highlight
    }
  }
}
