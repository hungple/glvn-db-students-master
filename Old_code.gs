////////////////////////////////////////////////////////////////////////////////////////
//
//     OLD - Nolonger needed
//
////////////////////////////////////////////////////////////////////////////////////////
/*
function readRow(className) {
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();

  var rowsDeleted = 0;
  for (var i = 0; i <= numRows - 1; i++) {
    var row = values[i];
    if (row[2] == 0 || row[2] == '') {
      //sheet.deleteRow((parseInt(i)+1) - rowsDeleted);
      //rowsDeleted++;
    }
  }  
};


function incrementCellValuesByOne() {
  // Increments the values in all the cells in the active range (i.e., selected cells).
  // Numbers increase by one, text strings get a "1" appended.
  // Cells that contain a formula are ignored.
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeRange = ss.getActiveRange();
  
  var cell, cellValue, cellFormula;
  
  // iterate through all cells in the active range
  for (var cellRow = 1; cellRow <= activeRange.getHeight(); cellRow++) {
    for (var cellColumn = 1; cellColumn <= activeRange.getWidth(); cellColumn++) {
      cell = activeRange.getCell(cellRow, cellColumn);
      cellFormula = cell.getFormula();
      
      // if not a formula, increment numbers by one, or add "1" to text strings
      // if the leftmost character is "=", it contains a formula and is ignored
      // otherwise, the cell contains a constant and is safe to increment
      // does not work correctly with cells that start with '=
      if (cellFormula[0] != "=") {
        cellValue = cell.getValue();
        cell.setValue(cellValue + 1);
      }
    }
  }
};

function getSheetName() {
  return SpreadsheetApp.getActiveSpreadsheet().getName();
}

function classTotal(name) {
  var rowCount = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name).getLastRow() - 1;
  return rowCount;
}

function getClassTotal(className) {
  
  var grades, names;
  
  var ss =SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students");
  if(className.charAt(0) == 'G') {
    grades = ss.getRange("G2:G").getValues();
    names = ss.getRange("H2:H").getValues();
  }
  else {
    grades = ss.getRange("I2:I").getValues();
    names = ss.getRange("J2:J").getValues();
  }
  var rowCount = 0;
  
  var row, len;
  len = grades.length;
  
  // iterate through all cells
  for (row = 0; row < len; row++) {
      if (grades[row] == className.charAt(2) && names[row] == className.charAt(3)) {
         rowCount = rowCount + 1;
      }
  }
  
  return rowCount;
}

//return select A,B,C,E where F='Yes' and G=1 and H='C' order by B,C
function classQuery() {
  var prefix = "select A,B,D,E,N,O,Q where F='Yes'";
  var sheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  if(sheetName.charAt(0) == "G")
    return prefix + " and G=" + sheetName.charAt(2) + " and H='" + sheetName.charAt(3) + "' order by B,C";
  else
    return prefix + " and I=" + sheetName.charAt(2) + " and J='" + sheetName.charAt(3) + "' order by B,C";
}
*/

/*
function createClasses() {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  
  var response = ui.alert(
      'Warning!!!',
      'Do you want to create new classes?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response == ui.Button.YES) {
    createClassesImpl();
  }
}


function createClassesImpl() {
  var glClassTemplateId = getStr("GL_CLASS_TEMPLATE_ID");
  var vnClassTemplateId = getStr("VN_CLASS_TEMPLATE_ID");

  createClassesImpl2("gl-classes", glClassTemplateId, 1);
  createClassesImpl2("vn-classes", vnClassTemplateId, 3);
}

function createClassesImpl2(sheetName, templateId, tokenNumber) {
  
  var clsNameCol          = 2;
  var gmailCol            = 6;
  var actionCol           = 7;
  var clsFolderIdCol      = 1;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  var range = sheet.getRange(2, 1, 20, 15); //row, col, numRows, numCols

  var clsName, action, clsFolder;
  
  // iterate through all cells in the range
  for (var cellRow = 1; cellRow <= range.getHeight(); cellRow++) {
    clsName = range.getCell(cellRow, clsNameCol).getValue();
    if( clsName == "")
      break;

    //gmail = range.getCell(cellRow, gmailCol).getValue().trim();
    action = range.getCell(cellRow, actionCol).getValue();
    clsFolder = DriveApp.getFolderById(range.getCell(cellRow, clsFolderIdCol).getValue());
    
    if(action == 'x') {
      Logger.log(clsName);

      // Make a copy and save it into the class folder
      //var file = DriveApp.getFileById(templateId);
      //var newFile = file.makeCopy(clsName, clsFolder);

      var template = DriveApp.getFileById(templateId);
      var classSht = getReportFolderId(clsName, clsFolder);

      /////////////////////////////////////////////////////////////////////////////
      // Open the new spreadsheet and setup basic functions
      /////////////////////////////////////////////////////////////////////////////
      var ss2 = SpreadsheetApp.openById(newFile.getId());
      ss2.getSheets()[0].getRange("A1:A1").getCell(1, 1).setValue(clsName);
      var temp = "=IMPORTRANGE(\"" + ss.getId() + "\",A1&\"!B1:I50\")";
      ss2.getSheets()[0].getRange("A2:A2").getCell(1, 1).setValue(temp);
      temp = "=IMPORTRANGE(\"" + ss.getId() + "\",\"" + sheetName + "\"&\"!E\"&(2*mid(A1,3,1)+if(right(A1,1)=\"A\",0,1)))";
      ss2.getSheets()[0].getRange("B1:B1").getCell(1, 1).setValue(temp);
      temp = "=IMPORTRANGE(\"" + ss.getId() + "\",\"calendar!B1:S1\")";
      ss2.getSheets()[1].getRange("F2:F2").getCell(1, 1).setValue(temp);      
      temp = "=IMPORTRANGE(\"" + ss.getId() + "\",\"calendar!B2:S2\")";
      ss2.getSheets()[2].getRange("F2:F2").getCell(1, 1).setValue(temp);      
      ss2.getSheets()[3].getRange("H3:H3").getCell(1, 1).setValue((cellRow/10+tokenNumber));
      Logger.log("Basic updated.");
      
      // Save report card folder id for each each class
      var reportFolderId = getReportFolderId(clsName, clsFolder);
      ss2.getSheets()[5].getRange("B3:B3").getCell(1, 1).setValue(reportFolderId);
      Logger.log("Update report card folder id: " + reportFolderId);
     
      /////////////////////////////////////////////////////////////////////////////
      // Save new class spreadsheet id into the class (ex: GL1A) sheet in the masters book
      /////////////////////////////////////////////////////////////////////////////
      var clsSheet = ss.getSheetByName(clsName);
      var clsRange = clsSheet.getRange(2, 1, 20, 15); //row, col, numRows, numCols
      var tstr = "=IMPORTRANGE(\"" + newFile.getId() + "\",\"Grades!F3:F50\")";
      clsRange.getCell(1, 14).setValue(tstr);
      Logger.log("Spreadsheet id is saved in " + clsName + " sheet of the master book.");
      
      /////////////////////////////////////////////////////////////////////////////
      // Save new class spreadsheet id into the gl-honor-roll sheet in the masters book
      /////////////////////////////////////////////////////////////////////////////
      var maxHonorRollEachClass = parseInt(getStr("MAX_HONOR_ROLL_EACH_CLASS"));
      // Only pull this maxHonorRollEachClass rows
      var imptStr = "=IMPORTRANGE(\"" + newFile.getId() + "\",\"honor-roll!B3:F" + (3+maxHonorRollEachClass-1) + "\")";
      //Logger.log(sheetName + " - " + imptStr);
      var hrSheet = ss.getSheetByName(sheetName.substring(0, 3)+"honor-roll");
      var hrRange = hrSheet.getRange(2, 1, 170, 15); //row, col, numRows, numCols
      var hrCell  = hrRange.getCell(1+((cellRow-1)*5), 2);
      hrCell.setValue(imptStr);
      Logger.log("Spreadsheet id is saved in the honor roll sheet.");
    }
  }
}


function getReportFolderId(clsName, clsFolder) {
  var files = clsFolder.getFilesByName(clsName);
  if (files.hasNext()) {
    var file = files.next();
    return file.getId();
  }
  return "";
}

function getReportFolderId(clsName, clsFolder) {
  var folders = clsFolder.getFoldersByName(clsName + "-report-cards");
  if (folders.hasNext()) {
    var folder = folders.next();
    return folder.getId();
  }
  return "";
}
*/