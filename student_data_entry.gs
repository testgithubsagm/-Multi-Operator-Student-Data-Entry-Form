function handleDataEntry() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var inputSheet = ss.getSheetByName("Input");
  var dataSheet = ss.getSheetByName("Data");

  if (!inputSheet || !dataSheet) {
    SpreadsheetApp.getUi().alert("Error: 'Input' or 'Data' sheet not found.");
    return;
  }

  // Get input values
  var rollNumber = inputSheet.getRange("D5").getValue().toString().trim();
  var name = inputSheet.getRange("D7").getValue().toString().trim();
  var grade = inputSheet.getRange("D9").getValue().toString().trim();
  var gender = inputSheet.getRange("D11").getValue().toString().trim();
  var admissionDate = inputSheet.getRange("D13").getValue();
  var remarks = inputSheet.getRange("D15").getValue().toString().trim();
  var operatorName = inputSheet.getRange("F3").getValue().toString().trim();
  var entryDate = new Date(); // Current Date

  // Ensure all required fields are filled
  if (!rollNumber || !name || !grade || !gender || !admissionDate || !remarks || !operatorName) {
    SpreadsheetApp.getUi().alert("Error: Please fill all required fields before submitting.");
    return;
  }

  // Check if Roll Number already exists
  var data = dataSheet.getDataRange().getValues();
  var rollColumnIndex = 0; // Assuming Roll Number is in the first column (A)

  for (var i = 1; i < data.length; i++) {
    if (data[i][rollColumnIndex].toString().trim() === rollNumber) {
      // If Roll Number exists, bring data back to input sheet
      inputSheet.getRange("D7").setValue(data[i][1]);  // Name
      inputSheet.getRange("D9").setValue(data[i][2]);  // Grade
      inputSheet.getRange("D11").setValue(data[i][3]); // Gender
      inputSheet.getRange("D13").setValue(data[i][4]); // Admission Date
      inputSheet.getRange("D15").setValue(data[i][5]); // Remarks
      SpreadsheetApp.getUi().alert("Roll number exists. Data retrieved.");
      return;
    }
  }

  // Append new row
  var nextRow = dataSheet.getLastRow() + 1;
  dataSheet.getRange(nextRow, 1, 1, 8).setValues([
    [rollNumber, name, grade, gender, admissionDate, remarks, operatorName, entryDate]
  ]);

  // Confirmation message
  SpreadsheetApp.getUi().alert("New data entered.");
}


function searchRollNumber() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var inputSheet = ss.getSheetByName("Input");
  var dataSheet = ss.getSheetByName("Data");

  if (!inputSheet || !dataSheet) {
    SpreadsheetApp.getUi().alert("Error: 'Input' or 'Data' sheet not found.");
    return;
  }

  var rollNumber = inputSheet.getRange("D5").getValue().toString().trim();
  
  if (!rollNumber) {
    SpreadsheetApp.getUi().alert("Enter a roll number.");
    return;
  }

  var data = dataSheet.getDataRange().getValues();
  var rollColumnIndex = 0; // Assuming Roll Number is in column A

  for (var i = 1; i < data.length; i++) {
    if (data[i][rollColumnIndex].toString().trim() === rollNumber) {
      // Populate the Input Sheet
      inputSheet.getRange("D7").setValue(data[i][1]);  // Name
      inputSheet.getRange("D9").setValue(data[i][2]);  // Grade
      inputSheet.getRange("D11").setValue(data[i][3]); // Gender
      inputSheet.getRange("D13").setValue(data[i][4]); // Date of Admission
      inputSheet.getRange("D15").setValue(data[i][5]); // Remarks
      SpreadsheetApp.getUi().alert("Roll number found. Data retrieved.");
      return;
    }
  }

  SpreadsheetApp.getUi().alert("Roll number not found.");
}



function editEntry() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var inputSheet = ss.getSheetByName("Input");
  var dataSheet = ss.getSheetByName("Data");

  if (!inputSheet || !dataSheet) {
    SpreadsheetApp.getUi().alert("Error: 'Input' or 'Data' sheet not found.");
    return;
  }

  var rollNumber = inputSheet.getRange("D5").getValue().toString().trim();
  
  if (!rollNumber) {
    SpreadsheetApp.getUi().alert("Enter a roll number.");
    return;
  }

  var data = dataSheet.getDataRange().getValues();
  var rollColumnIndex = 0; // Assuming Roll Number is in column A
  var operatorColumnIndex = 6; // Assuming Operator Name is in column G (after Remarks)

  for (var i = 1; i < data.length; i++) {
    if (data[i][rollColumnIndex].toString().trim() === rollNumber) {
      // Update existing row with new data, including the operator name
      dataSheet.getRange(i + 1, 2, 1, 5).setValues([[
        inputSheet.getRange("D7").getValue(),  // Name
        inputSheet.getRange("D9").getValue(),  // Grade
        inputSheet.getRange("D11").getValue(), // Gender
        inputSheet.getRange("D13").getValue(), // Date of Admission
        inputSheet.getRange("D15").getValue()  // Remarks
      ]]);

      // ✅ Update Operator Name
      var operatorName = inputSheet.getRange("F3").getValue();
      dataSheet.getRange(i + 1, operatorColumnIndex + 1).setValue(operatorName);

      SpreadsheetApp.getUi().alert("Entry updated.");
      return;
    }
  }

  SpreadsheetApp.getUi().alert("Roll number not found.");
}


function clearFields() {
  var inputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Input");
  
  if (!inputSheet) {
    SpreadsheetApp.getUi().alert("Error: 'Input' sheet not found.");
    return;
  }

  inputSheet.getRangeList(["D5", "D7", "D9", "D11", "D13", "D15"]).clearContent();
  SpreadsheetApp.getUi().alert("Fields cleared.");
}












