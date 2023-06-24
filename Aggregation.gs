function copyData() {
  // Stage 1: Define parameters
  // Target Table
  var targetTableId = '1olsYDb9PI4ZgM37_03CRnH7ms8opywmzAtGzqK9mFXs';
  var targetSheetName = 'Sheet1';

  // Source Sheet 1: Auto fit data
  var sourceSheet1Id = '1ldqJb5I15kUMxX6MLlq1PNRA6LT3BZWEZMkf67vxPS8';
  var sourceSheet1Name = 'Sheet2';

  // Source Sheet 2: Interested Daily
  var sourceSheet2Id = '1DjJwtLzOIsxpOFmY_uPvOpCXFXRnpQQOq2bbhgKcMb8';
  var sourceSheet2Name = 'Output';

  // Source Sheet 3: Dashboard Form
  var sourceSheet3Id = '1nyhu6qrJCd-EqXHSKd7ZRT8YBXmo8PE8f4mEzky3ddU';
  var sourceSheet3Name = 'Form';
  
  // Source Sheet 4: New position form
  var sourceSheet4Id = '1puLaB-YYTlMF2Q-AzbzZwC6f5DjRUJZwnTFLl2ldra8';
  var sourceSheet4Name = 'P_ID';

  // Stage 2: Create target table
  var targetTable = SpreadsheetApp.openById(targetTableId).getSheetByName(targetSheetName);
  targetTable.getRange(1, 1, 1, 10).setValues([['candidate_id', 'position_id', 'fitPercentage', 'Min Date', 'Remove', 'Remove C', 'Remove P', 'Name (HR)', 'Phone (HR)', 'Mail (HR)']]);

  // Stage 3: Import fit data
  var sourceSheet1 = SpreadsheetApp.openById(sourceSheet1Id).getSheetByName(sourceSheet1Name);
  var sourceData1 = sourceSheet1.getRange('A:C').getValues();

  // Remove the header row from sourceData1
  sourceData1.shift();

  targetTable.getRange(2, 1, sourceData1.length, sourceData1[0].length).setValues(sourceData1);

  // Stage 4: Import Interested
  var sourceSheet2 = SpreadsheetApp.openById(sourceSheet2Id).getSheetByName(sourceSheet2Name);
  var sourceData2 = sourceSheet2.getDataRange().getValues();

  // Remove the header row from sourceData2
  sourceData2.shift();

  for (var i = 0; i < sourceData2.length; i++) {
    var candidateId = sourceData2[i][0];
    var positionId = sourceData2[i][1];
    var minDate = sourceData2[i][2];
    var targetData = targetTable.getDataRange().getValues();
    var found = false;
    
    for (var j = 1; j < targetData.length; j++) {
      if (targetData[j][0] == candidateId && targetData[j][1] == positionId) {
        targetTable.getRange(j + 1, 4).setValue(minDate);  // if found match, update Min Date
        found = true;
        break;
      }
    }

    if (!found) {
      targetTable.appendRow([candidateId, positionId, '', minDate, '', '', '']);  // if not found, append new row
    }
  }

  // Stage 5: Import Remove C_P
  var sourceSheet3 = SpreadsheetApp.openById(sourceSheet3Id).getSheetByName(sourceSheet3Name);
  var sourceData3 = sourceSheet3.getDataRange().getValues();

  // Remove the header row from sourceData3
  sourceData3.shift();

  for (var i = 0; i < sourceData3.length; i++) {
    var candidateId = sourceData3[i][0];
    var positionId = sourceData3[i][1];
    var removeCP = sourceData3[i][2];
    var targetData = targetTable.getDataRange().getValues();
    var found = false;
    
    for (var j = 1; j < targetData.length; j++) {
      if (targetData[j][0] == candidateId && targetData[j][1] == positionId) {
        targetTable.getRange(j + 1, 5).setValue(removeCP);  // if found match, update Remove C_P
        found = true;
        break;
      }
    }

    if (!found) {
      targetTable.appendRow([candidateId, positionId, '', '', removeCP, '', '']);  // if not found, append new row
    }
  }
  
  // Stage 6: Import Remove C
  for (var i = 0; i < sourceData3.length; i++) {
    var candidateId = sourceData3[i][0];
    var removeC = sourceData3[i][2];
    var targetData = targetTable.getDataRange().getValues();

    // Scan through the target data, flag if match is found
    for (var j = 1; j < targetData.length; j++) {
      if (targetData[j][0] == candidateId) {
        targetTable.getRange(j + 1, 6).setValue(removeC);  // if found match, update Remove C
      }
    }

    // Check if candidateId does not exist in the targetData
    var found = targetData.some(function(row) {
      return row[0] == candidateId;
    });

    if (!found) {
      targetTable.appendRow([candidateId, '', '', '', '', removeC, '']);  // if not found, append new row
    }
  }

  // Stage 7: Import Remove P
  for (var i = 0; i < sourceData3.length; i++) {
    var positionId = sourceData3[i][1];
    var removeP = sourceData3[i][2];
    var targetData = targetTable.getDataRange().getValues();

    // Scan through the target data, flag if match is found
    for (var j = 1; j < targetData.length; j++) {
      if (targetData[j][1] == positionId) {
        targetTable.getRange(j + 1, 7).setValue(removeP);  // if found match, update Remove P
      }
    }

    // Check if positionId does not exist in the targetData
    var found = targetData.some(function(row) {
      return row[1] == positionId;
    });

    if (!found) {
      targetTable.appendRow(['', positionId, '', '', '', '', removeP]);  // if not found, append new row
    }
  }

  // Stage 8: HR Details
  var sourceSheet4 = SpreadsheetApp.openById(sourceSheet4Id).getSheetByName(sourceSheet4Name);
  var sourceData4 = sourceSheet4.getDataRange().getValues().slice(1); // Exclude headers

  var targetData = targetTable.getDataRange().getValues();
  var positionIdToRowIndex = {};

  // Map positionId to row index
  for (var j = 0; j < targetData.length; j++) {
    positionIdToRowIndex[targetData[j][1]] = j;
  }

  for (var i = 0; i < sourceData4.length; i++) {
    var positionId = sourceData4[i][0];
    var nameHR = sourceData4[i][1];
    var phoneHR = sourceData4[i][2];
    var mailHR = sourceData4[i][3];

    if (positionId in positionIdToRowIndex) {
      // If found match, update Name (HR), Phone (HR), and Mail (HR)
      var rowIndex = positionIdToRowIndex[positionId];
      targetTable.getRange(rowIndex + 1, 8).setValue(nameHR);
      targetTable.getRange(rowIndex + 1, 9).setValue(phoneHR);
      targetTable.getRange(rowIndex + 1, 10).setValue(mailHR);
    } else {
      // If not found, append new row
      targetTable.appendRow(['', positionId, '', '', '', '', '', nameHR, phoneHR, mailHR]);
      // Update our map with the new row
      positionIdToRowIndex[positionId] = targetData.length;
      targetData.push(['', positionId, '', '', '', '', '', nameHR, phoneHR, mailHR]);
    }
  }
}

