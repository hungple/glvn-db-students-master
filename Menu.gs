/////////////////////////////////////////////////////////////////////////////////////////////////////
//
// GLVN menu item
//
/////////////////////////////////////////////////////////////////////////////////////////////////////

/**
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [
    // DO NOT DELETE THESE TWO FUNCTIONS
    //{ 
    //  name : "Create classes (root only)",
    //  functionName : "createClasses"
    //},    
    //{ 
    //  name : "Update class sheets in this spreadsheet (root only)",
    //  functionName : "updateSheetsInThisSpreadSheet"
    //},    
    { 
      name : "Update class spreadSheets (root only)",
      functionName : "updateClassSpreadSheets"
    },    
    //{ 
    //  name : "Update class spreadSheet using template spreadSheet(root only)",
    //  functionName : "updateClassSpreadSheetUsingTemplateSpreadSheet"
    //},    
    //{
    //name : "00 - Fill in service dates (August/early September)",
    //functionName : "fillServiceDates"
    //},
    {
      name : "01 - Share classes (September/Octorber)",
      functionName : "shareClasses"
    },
    {
      name : "02 - Update student glFinalPoint and vnFinalPoint (end of May/early June)",
      functionName : "updateFinalPoints"
    },
    {
      name : "03 - Update First Communion date and location (end of May/early June)",
      functionName : "updateFCommunionInfo"
    },
    {
      name : "04 - Update Confirmation date and location (end of May/early June)",
      functionName : "updateConfirmationInfo"
    },
    {
      name : "05 - Un-share classes (June)",
      functionName : "unShareClasses"
    },
    {
      name : "06 - Save students into former-students folder (June)",
      functionName : "saveFormerStudents"
    },
    {
      name : "07 - Update student classes for new registration (end of June/July)",
      functionName : "updateClassesForNewReg"
    }

    ];
  sheet.addMenu("GLVN", entries);
};


function getStr(key) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("admin");
  var range = sheet.getRange(2, 1, 20, 2); //row, col, numRows, numCols

  // iterate through all cells in the range
  for (var cellRow = 1; cellRow <= range.getHeight(); cellRow++) {
    var varName = range.getCell(cellRow, 1).getValue();
    if( varName == key){
      return range.getCell(cellRow, 2).getValue();
    }
  }
  return "";
}


/////////////////////////////////////////////////////////////////////////////////////////////////////
//
// DO NOT DELETE THIS FUNCTION
// Create classes
// This function is for setting up database for the first time only
//
/////////////////////////////////////////////////////////////////////////////////////////////////////
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

  var admins = getStr("ADMIN_IDS").split(",");
  for (var i = 0; i < admins.length; i++) {
    admins[i] = admins[i].trim();
  }
  
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
      
      // Unsharing GL1A and rename it to "bk"
      
      var files = clsFolder.getFilesByName(clsName); // open spreadsheet GL1A
      while (files.hasNext()) {
        var file = files.next();
        var spreadSheet = SpreadsheetApp.openById(file.getId());
        Logger.log("Spreadsheet: " + spreadSheet.getName());
        try {
          //remove editors
          var editors = spreadSheet.getEditors();
          for (var j = 0; j < editors.length; j++) {
            if(isNotAdmin(admins, editors[j].getEmail())){
              spreadSheet.removeEditor(editors[j].getEmail());
              Logger.log("Editor removed: " + editors[j].getEmail());              
            }
          }             
        }
        catch(e) {
          Logger.log(e); //ignore error
        }
        
        // rename to "bk"
        file.setName("bk");
      }

      /////////////////////////////////////////////////////////////////////////////
      // Make a copy and save it into the class folder
      /////////////////////////////////////////////////////////////////////////////
      var file = DriveApp.getFileById(templateId);
      var newFile = file.makeCopy(clsName, clsFolder);

      /////////////////////////////////////////////////////////////////////////////
      // Open the new spreadsheet and setup basic functions
      /////////////////////////////////////////////////////////////////////////////
      var newss = SpreadsheetApp.openById(newFile.getId());
      newss.getSheetByName("contacts").getRange("A1:A1").getCell(1, 1).setValue(clsName);

      var temp = "=IMPORTRANGE(\"" + ss.getId() + "\",A1&\"!B1:I50\")";
      newss.getSheetByName("contacts").getRange("A2:A2").getCell(1, 1).setValue(temp);
      temp = "=IMPORTRANGE(\"" + ss.getId() + "\",\"" + sheetName + "\"&\"!E\"&(2*mid(A1,3,1)+if(right(A1,1)=\"A\",0,1)))";
      newss.getSheetByName("contacts").getRange("B1:B1").getCell(1, 1).setValue(temp);
      
      temp = "=IMPORTRANGE(\"" + ss.getId() + "\",\"calendar!B1:S1\")";
      newss.getSheetByName("attendance-HK1").getRange("F2:F2").getCell(1, 1).setValue(temp);      
      
      temp = "=IMPORTRANGE(\"" + ss.getId() + "\",\"calendar!B2:S2\")";
      newss.getSheetByName("attendance-HK2").getRange("F2:F2").getCell(1, 1).setValue(temp);      
      
      newss.getSheetByName("grades").getRange("H3:H3").getCell(1, 1).setValue((cellRow/10+tokenNumber));
      Logger.log("Basic updated.");
      
      // Save report card folder id for each each class
      var reportFolderId = getReportFolderId(clsName, clsFolder);
      newss.getSheetByName("admin").getRange("B3:B3").getCell(1, 1).setValue(reportFolderId);
      Logger.log("Update report card folder id: " + reportFolderId);
     
      /////////////////////////////////////////////////////////////////////////////
      // Save new class spreadsheet id into the class (ex: GL1A) sheet in the masters book
      /////////////////////////////////////////////////////////////////////////////
      var tstr = "=IMPORTRANGE(\"" + newFile.getId() + "\",\"Grades!F3:F50\")";
      ss.getSheetByName(clsName).getRange("N2:N2").getCell(1, 1).setValue(tstr);
      Logger.log("Spreadsheet id is saved in " + clsName + " sheet of the master book.");
      
      /////////////////////////////////////////////////////////////////////////////
      // Save new class spreadsheet id into the gl/vn-honor-roll sheet in the masters book
      /////////////////////////////////////////////////////////////////////////////
      var maxHonorRollEachClass = parseInt(getStr("MAX_HONOR_ROLL_EACH_CLASS"));
      // Only pull this maxHonorRollEachClass rows
      var imptStr = "=IMPORTRANGE(\"" + newFile.getId() + "\",\"honor-roll!B3:F" + (3+maxHonorRollEachClass-1) + "\")";
      var hrSheet = ss.getSheetByName(sheetName.substring(0, 3)+"honor-roll"); // gl-honor-roll or vn-honor-roll
      var hrRange = hrSheet.getRange(2, 1, 170, 15); //row, col, numRows, numCols
      var hrCell  = hrRange.getCell(1+((cellRow-1)*5), 2);
      hrCell.setValue(imptStr);
      Logger.log("Spreadsheet id is saved in the honor roll sheet.");
    }
  }
}


function getClassSpreadsheetId(clsName, clsFolder) {
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




/////////////////////////////////////////////////////////////////////////////////////////////////////
//
// DO NOT DELETE THIS FUNCTION
// Update classes
// This function can be used for updating individual cell in each class sheet
//
/////////////////////////////////////////////////////////////////////////////////////////////////////
function updateSheetsInThisSpreadSheet() {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  
  var response = ui.alert(
      'Warning!!!',
      'Do you want to updateSheetsInThisSpreadSheet?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response == ui.Button.YES) {
    updateSheetsInThisSpreadSheetImpl();
  }
}


function updateSheetsInThisSpreadSheetImpl() {
  updateSheetsInThisSpreadSheetImpl2("gl-classes");
  updateSheetsInThisSpreadSheetImpl2("vn-classes");
}


function updateSheetsInThisSpreadSheetImpl2(sheetName) {
  
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
      
      updateSheetsInThisSpreadSheet_ClassSheet(ss.getSheetByName(clsName), clsName);
    }
  }
}


function updateSheetsInThisSpreadSheet_ClassSheet(sheet, clsName) {

  var newValue
  if(clsName.substr(0,2) == "GL") { // GL class sheet
    newValue = '=query(students!1:700, "select A,B,C,D,E,R,O,P,Q,I,J,K,AD where " & if(left(A1,1)="G","G","I") & "=" & mid(A1,3,1) & " and " & if(left(A1,1)="G","H","J") & "=\'" & right(A1,1) & "\' order by C,E")'
  }
  else { // VN class sheet
    newValue = '=query(students!1:700, "select A,B,C,D,E,R,O,P,Q,G,H,K,AD where " & if(left(A1,1)="G","G","I") & "=" & mid(A1,3,1) & " and " & if(left(A1,1)="G","H","J") & "=\'" & right(A1,1) & "\' order by C,E")'
  }
  sheet.getRange("B1:B1").getCell(1, 1).setValue(newValue);
  

  newValue = 'TotalPoints'
  sheet.getRange("O1:O1").getCell(1, 1).setValue(newValue);
  
  newValue = '=CONCATENATE(COUNTIFS(O2:O52, "0")," | ", MIN(O2:O52)," - ", MAX(O2:O52), " | ", COUNTIFS(O2:O52, ">89.99"), ":", COUNTIFS(O2:O52, ">79.99")-COUNTIFS(O2:O52, ">89.99"), ":", COUNTIFS(O2:O52, ">69.99")-COUNTIFS(O2:O52, ">79.99"), ":", COUNTIFS(O2:O52, "<=69.99"))'
  sheet.getRange("P1:P1").getCell(1, 1).setValue(newValue);
}




/////////////////////////////////////////////////////////////////////////////////////////////////////
//
// DO NOT DELETE THIS FUNCTION
// Update classes
// This function can be used for updating individual cell in each class sheet
//
/////////////////////////////////////////////////////////////////////////////////////////////////////
function updateClassSpreadSheets() {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  
  var response = ui.alert(
      'Warning!!!',
      'Do you want to updateClassSpreadSheets?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response == ui.Button.YES) {
    updateClassSpreadSheetsImpl();
  }
}


function updateClassSpreadSheetsImpl() {
  updateClassSpreadSheetsImpl2("gl-classes");
  updateClassSpreadSheetsImpl2("vn-classes");
}


function updateClassSpreadSheetsImpl2(sheetName) {
  
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
      
      // Open target class spreadsheet
      var tss = SpreadsheetApp.openById(getClassSpreadsheetId(clsName, clsFolder));

      // Update `contacts` sheet in each class spreadsheet
      updateClassSpreadSheets_contactsSheet(tss, clsName, ss.getId());

      // Update `attendance_HK1` sheet in each class spreadsheet
      updateClassSpreadSheets_attendanceSheet(tss, "HK1");

      // Update `attendance_HK2` sheet in each class spreadsheet
      updateClassSpreadSheets_attendanceSheet(tss, "HK2");

      // Update `grades` sheet in each class spreadsheet
      updateClassSpreadSheets_gradesSheet(tss, clsName);

    }
  }
}

function updateClassSpreadSheets_contactsSheet(tss, clsName, studentMasterSpreadsheetId) {

  var sn = 'contacts';
      
  // sheet
  var sheet = tss.getSheetByName(sn);
  
  // insert 1 column after column 3 in this sheet
  sheet.insertColumns(3, 1); 
  
  var newValue = '=IMPORTRANGE("' + studentMasterSpreadsheetId + '",A1&"!B1:J52")'
  sheet.getRange("A2:A2").getCell(1, 1).setValue(newValue);
}

function updateClassSpreadSheets_attendanceSheet(tss, hocKy) {

  var sn = 'attendance-' + hocKy;
      
  // sheet
  var sheet = tss.getSheetByName(sn);
  
  var newValue = '=query(contacts!2:52, "select C,E,G,H")'
  sheet.getRange("B2:B2").getCell(1, 1).setValue(newValue);
}

function updateClassSpreadSheets_gradesSheet(tss) {

  var sn = 'grades';
      
  // sheet
  var sheet = tss.getSheetByName(sn);
  
  var newValue = '=query(contacts!2:60, "select A,B,C,E,I")'
  sheet.getRange("A2:A2").getCell(1, 1).setValue(newValue);
}




/////////////////////////////////////////////////////////////////////////////////////////////////////
//
// DO NOT DELETE THIS FUNCTION
// Update classes
// This function can be used for updating classes if teachers have not updated their classes
//
/////////////////////////////////////////////////////////////////////////////////////////////////////
function updateClassSpreadSheetUsingTemplateSpreadSheet() {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  
  var response = ui.alert(
      'Warning!!!',
      'Do you want to updateClassSpreadSheetUsingTemplateSpreadSheet?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response == ui.Button.YES) {
    updateClassSpreadSheetUsingTemplateSpreadSheetImpl();
  }
}


function updateClassSpreadSheetUsingTemplateSpreadSheetImpl() {
  var glClassTemplateId = getStr("GL_CLASS_TEMPLATE_ID");
  var vnClassTemplateId = getStr("VN_CLASS_TEMPLATE_ID");

  updateClassSpreadSheetUsingTemplateSpreadSheetImpl2("gl-classes", glClassTemplateId, 1);
  updateClassSpreadSheetUsingTemplateSpreadSheetImpl2("vn-classes", vnClassTemplateId, 3);
}


function updateClassSpreadSheetUsingTemplateSpreadSheetImpl2(sheetName, templateId, tokenNumber) {
  
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

      // Open template spreadsheet
      var tss = SpreadsheetApp.openById(templateId);
      
      // Open class spreadsheet
      var css = SpreadsheetApp.openById(getClassSpreadsheetId(clsName, clsFolder));

      // Update `grades` sheet
      updateClassSpreadSheetUsingTemplateSpreadSheet_gradesSheet(tss, css, cellRow, tokenNumber);

      // Update `attendance_HK1` sheet
      updateClassSpreadSheetUsingTemplateSpreadSheet_attendanceSheet(tss, css);
    }
  }
}


function updateClassSpreadSheetUsingTemplateSpreadSheet_gradesSheet(tss, css, cellRow, tokenNumber) {

  var sn = 'grades';
      
  // source sheet
  var ss = tss.getSheetByName(sn);
  
  // set the token for verifying
  ss.getRange("H3:H3").getCell(1, 1).setValue((cellRow/10+tokenNumber));
  Logger.log((cellRow/10+tokenNumber));

  // Get full range of data
  var SRange = ss.getDataRange();

  // get A1 notation identifying the range
  var A1Range = SRange.getA1Notation();

  // get the data values in range
  var SData = SRange.getValues();
      
  // target sheet
  var ts = css.getSheetByName(sn); 

  // Clear the Google Sheet before copy
  ts.clear({contentsOnly: true});

  // set the target range to the values of the source data
  ts.getRange(A1Range).setValues(SData);
     
}


function updateClassSpreadSheetUsingTemplateSpreadSheet_attendanceSheet(tss, css) {

  var sn = 'attendance-HK1';
      
  // source sheet
  var ss = tss.getSheetByName(sn);
  
  // Get full range of data
  var SRange = ss.getDataRange();

  // get A1 notation identifying the range
  var A1Range = SRange.getA1Notation();

  // get the data values in range
  var SData = SRange.getValues();
      
  // target sheet
  var ts = css.getSheetByName(sn); 

  // Clear the Google Sheet before copy
  ts.clear({contentsOnly: true});

  // set the target range to the values of the source data
  ts.getRange(A1Range).setValues(SData);
     
}



/////////////////////////////////////////////////////////////////////////////////////////////////////
//
// 00 - Fill in the service dates
//
/////////////////////////////////////////////////////////////////////////////////////////////////////
function fillServiceDates() {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  
  var response = ui.alert(
      'Warning!!!',
      'Do you want to create or update service dates?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response == ui.Button.YES) {
    fillServiceDatesImpl();
  }
}

function fillServiceDatesImpl() {
  var idCol       = 1;  
  var glLevelCol  = 8; 
  var glNameCol   = 9; 
  var vnLevelCol  = 10; 
  var vnNameCol   = 11; 
  var isRegCol    = 12; 
  
  var serviceDateCol = 21;  

  var startCol    = 34;  

  var numberOfDates  = 24; //number of service dates
  var calSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Calendar");
  var calRange = calSheet.getRange(5, 2, 1, numberOfDates); //row, col, numRows, numCols
  
  ////////////////////////////////////////////////////////////////////////////////////
  var sheet        = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Std_VGz6v3");
  var range        = sheet.getRange(1, 1, 700, 34); //row, col, numRows, numCols 
  var rowStartCell = range.getCell(2, startCol);
  var rowStart     = rowStartCell.getValue()+1;
  var index1Cell   = range.getCell(3, startCol);
  var index1       = index1Cell.getValue();
  var index2Cell   = range.getCell(4, startCol);
  var index2       = index2Cell.getValue();
  
  var id, isReg, glLevel, glName, vnLevel, vnName; 
  // iterate through all cells in the range
  for (var cellRow = rowStart; ; cellRow++) {
    id = range.getCell(cellRow, idCol).getValue(); 

    if(id == "") break;
    
    isReg   = range.getCell(cellRow, isRegCol).getValue();
    glLevel = range.getCell(cellRow, glLevelCol).getValue();
    glName  = range.getCell(cellRow, glNameCol).getValue();
    vnLevel = range.getCell(cellRow, vnLevelCol).getValue();
    vnName  = range.getCell(cellRow, vnNameCol).getValue();
    
    if(isReg == "x") {
      var loc = getLocation(glLevel, glName);
      var dat;
      if(loc == 1) {
        dat = calRange.getCell(1,index1).getValue();
        index1 = index1 % numberOfDates + 1;
        index1Cell.setValue(index1);
      }
      else {
        dat = calRange.getCell(1,index2).getValue();
        index2 = index2 % numberOfDates + 1;
        index2Cell.setValue(index2);
      }
      
      range.getCell(cellRow, serviceDateCol).setValue(dat);
    }
   
    // update rowStart cell
    rowStartCell.setValue(cellRow);
  }
 
};
    
function getLocation(glLevel, glName) {
  if(glLevel > 5) {
    return 1;
  }
  else if(glName=='B') {
    return 1;
  }
  else {
    return 2;
  }
}



/////////////////////////////////////////////////////////////////////////////////////////////////////
//
// 01 - Share classes to the teachers
//
/////////////////////////////////////////////////////////////////////////////////////////////////////
function shareClasses() {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  
  var response = ui.alert(
      'Warning!!!',
      'Do you want to share classes to the teachers?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response == ui.Button.YES) {
    shareClassesImpl(true);
  }
}


function shareClassesImpl(isShared) {
  var glReportCardTemplateId = getStr("GL_REPORT_CARD_TEMPLATE_ID");
  var vnReportCardTemplateId = getStr("VN_REPORT_CARD_TEMPLATE_ID");
  shareClassesImpl2("gl-classes", glReportCardTemplateId, isShared);
  shareClassesImpl2("vn-classes", vnReportCardTemplateId, isShared);
}


function shareClassesImpl2(sheetName, reportFormId, isShared) {
  
  var clsFolderIdCol      = 1;
  var clsNameCol          = 2;
  var gmailCol            = 6;
  var actionCol           = 7;
  
  // Common files to share to the teachers
  //var commonFilesToShareFolderId  = getStr("DOC_SHARE_FOLDER_ID");
   
  var admins = getStr("ADMIN_IDS").split(",");
  for (var i = 0; i < admins.length; i++) {
    admins[i] = admins[i].trim();
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var range = sheet.getRange(2, 1, 19, 15); //row, col, numRows, numCols

  var clsName, gmails, clsFolder, action;
  
  // iterate through all cells in the range
  for (var cellRow = 1; cellRow <= range.getHeight(); cellRow++) {
    clsName = range.getCell(cellRow, clsNameCol).getValue();
    gmails = range.getCell(cellRow, gmailCol).getValue().trim();
    //if( gmails == "")
    //  break;
    
    action = range.getCell(cellRow, actionCol).getValue();
    clsFolder = DriveApp.getFolderById(range.getCell(cellRow, clsFolderIdCol).getValue());
    Logger.log(clsName + " Entry: '" + clsFolder.getName() + "'");
    
    if(action == "x") {
      
      var gmailArr = gmails.split(",");
      for (i = 0; i < gmailArr.length; i++) {
        gmailArr[i] = gmailArr[i].trim();
      }
      Logger.log(gmailArr);
      
      // report card template
      var doc = DocumentApp.openById(reportFormId);
      Logger.log(doc.getName()); 
      for (var i = 0; i < gmailArr.length; i++) {
        var gmail = gmailArr[i];
        try {
          //Only remove this gmail but don't remove other viewers
          doc.removeViewer(gmail);
          Logger.log("Viewer removed: " + gmail);
          
          if(isShared == true){
            doc.addViewer(gmail);
            Logger.log("Viewer added: " + gmail);
          }
        }
        catch(e) {Logger.log(e);} //ignore error
      }
      
      // report card folder
      var folders = clsFolder.getFolders();
      while (folders.hasNext()) {
        var folder = folders.next();
        Logger.log(folder.getName());
        
        try {
          //remove editors
          var editors = folder.getEditors();
          //Logger.log(editors);
          for (var j = 0; j < editors.length; j++) {
            if(isNotAdmin(admins, editors[j].getEmail())){
              folder.removeEditor(editors[j].getEmail());
              Logger.log("Editor removed: " + editors[j].getEmail());              
            }
          }
          
          // add new editor
          for (var i = 0; i < gmailArr.length; i++) {
            var gmail = gmailArr[i];
            if(isShared == true){
              folder.addEditor(gmail);
              Logger.log("Editor added: " + gmail);
            }
          }
        }
        catch(e) {Logger.log(e);} //ignore error
      } //while folder
    
      // Sharing files
      var files = clsFolder.getFiles();
      while (files.hasNext()) {
        var file = files.next();
        
        if(file.getName().length == 4) {
          // spreadsheet GL1A
          var spreadSheet = SpreadsheetApp.openById(file.getId());
          Logger.log("Spreadsheet: " + spreadSheet.getName());
          try {
            //remove editors
            var editors = spreadSheet.getEditors();
            for (var j = 0; j < editors.length; j++) {
              if(isNotAdmin(admins, editors[j].getEmail())){
                spreadSheet.removeEditor(editors[j].getEmail());
                Logger.log("Editor removed: " + editors[j].getEmail());              
              }
            }             
            
            // add new editor
            for (var i = 0; i < gmailArr.length; i++) {
              var gmail = gmailArr[i];
              if(isShared == true){
                spreadSheet.addEditor(gmail);
                Logger.log("Editor added: " + gmail);
              }
            }
          }
          catch(e) {Logger.log(e);} //ignore error
        }
        else if(file.getName().length > 4)  { // Share other docs only filename that is longer than 4 characters
          var doc = DocumentApp.openById(file.getId());
          Logger.log(doc.getName());
          try {
            //remove editors
            var editors = doc.getEditors();
            for (var j = 0; j < editors.length; j++) {
              if(isNotAdmin(admins, editors[j].getEmail())){
                doc.removeEditor(editors[j].getEmail());
                Logger.log("Editor removed: " + editors[j].getEmail());              
              }
            }             

            // add new editor
            for (var i = 0; i < gmailArr.length; i++) {
              var gmail = gmailArr[i];
              if(isShared == true){
                doc.addEditor(gmail);
                Logger.log("Editor added: " + gmail);
              }
            }
          }
          catch(e) {Logger.log(e);} //ignore error
        }
      } //while
    
      Logger.log(clsName + " - exit'");
    }
  }
};

function isNotAdmin(admins, gmail) {
  //Logger.log(admins);
  //Logger.log(gmail);
  for (var i = 0; i < admins.length; i++) {
    if(admins[i].toUpperCase() === gmail.toUpperCase()) {
      return false;
    }
  }
  return true;
}

/////////////////////////////////////////////////////////////////////////////////////////////////////
//
// 02 - Update Final Points
//
/////////////////////////////////////////////////////////////////////////////////////////////////////
function updateFinalPoints() {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  
  var response = ui.alert(
      'Warning!!!',
      'Do you want to update glFinalPoint and vnFinalPoint for students?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response == ui.Button.YES) {
    updateFinalPointImpl();
  }
}


function updateFinalPointImpl() {
  
  var idCol       = 1;  
  var glLevelCol  = 7; 
  var glNameCol   = 8; 
  var vnLevelCol  = 9; 
  var vnNameCol   = 10; 
  var isRegCol    = 11; 
  
  var glFinalGradeCol = 32;  
  var vnFinalGradeCol = 33; 
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Std_VGz6v3");
  var range = sheet.getRange(1, 1, 700, 34); //row, col, numRows, numCols
  var rowStartCell = sheet.getRange("AH1:AH1").getCell(1, 1); // <= Need to update column
  ////////////////////////////////////////////////////////////////////////////////////
  
  var id, isReg, glLevel, glName, vnLevel, vnName; 
  
  // iterate through all cells in the range
  for (var cellRow = rowStartCell.getValue(); ; cellRow++) {
    
    id = range.getCell(cellRow, idCol).getValue(); 
    if(id == "") break;
    
    isReg   = range.getCell(cellRow, isRegCol).getValue();
    glLevel = range.getCell(cellRow, glLevelCol).getValue();
    glName  = range.getCell(cellRow, glNameCol).getValue();
    vnLevel = range.getCell(cellRow, vnLevelCol).getValue();
    vnName  = range.getCell(cellRow, vnNameCol).getValue();
    
    if(isReg == "x") {
      if(glName != "") {
        range.getCell(cellRow, glFinalGradeCol).setValue(getFinalGrade("GL" + glLevel + glName, id));
      }
      
      if(vnName != "") {
        range.getCell(cellRow, vnFinalGradeCol).setValue(getFinalGrade("VN" + vnLevel + vnName, id));
      }
    }
    else {
      range.getCell(cellRow, glFinalGradeCol).setValue("");
      range.getCell(cellRow, vnFinalGradeCol).setValue("");
    }
   
    // update rowStart cell
    rowStartCell.setValue(cellRow);
  }
  
};


function getFinalGrade(className, sid) { // sheet GL1A, GL1B, GL2A...
  var idCol = 2; // col B
  var totalPointsCol = 15;
  
  //========================================================================
  
  var sheet =SpreadsheetApp.getActiveSpreadsheet().getSheetByName(className);
  if(sheet != null) {
    var range = sheet.getRange(2, 1, 60, 20); //row, col, numRows, numCols
  
    var idCell, totalPointsCell;
  
    // iterate through all cells in the range
    for (var cellRow = 1; cellRow <= range.getHeight(); cellRow++) {
      idCell = range.getCell(cellRow, idCol);
      totalPointsCell = range.getCell(cellRow, totalPointsCol);
      if(idCell.getValue() == sid) {
        return totalPointsCell.getValue();
      }
    }
  }
  return 0;
};


/////////////////////////////////////////////////////////////////////////////////////////////////////
//
// 03 - Update first communion date and location
//
/////////////////////////////////////////////////////////////////////////////////////////////////////
function updateFCommunionInfo() {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  
  var response = ui.alert(
      'Warning!!!',
      'Do you want to update communion information for students?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response == ui.Button.YES) {
    updateFCommunionInfoImpl();
  }
}

function updateFCommunionInfoImpl() {
  var idCol       = 1;
  var glLevelCol  = 7;
  var glNameCol   = 8;
  var isRegCol    = 11;
  
  var commDateCol     = 25;
  var commLocationCol = 26;
  
  var glFinalPointCol   = 32;  
  
  //////////////////////////// Get data from the Calendar sheet  //////////////////////////////
  var varSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Calendar");
  var varRange = varSheet.getRange(1, 1, 10, 11); //row, col, numRows, numCols 
  var commDate = varRange.getCell(3, 2).getValue(); //row, col
  var commLocation = "St. Maria Goretti, San Jose, CA"
  ////////////////////////////////////////////////////////////////////////////////////
  
  
  ////////////////////////////////////////////////////////////////////////////////////
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Std_VGz6v3");
  var range = sheet.getRange(1, 1, 700, 34); //row, col, numRows, numCols 
  var rowStartCell = range.getCell(2, 34);
  var rowStart = rowStartCell.getValue()+1;
  
  var id, isReg, glLevel, glName, glFinalPoint; 
  // iterate through all cells in the range
  for (var cellRow = rowStart; ; cellRow++) {
    id = range.getCell(cellRow, idCol).getValue(); 

    if(id == "") break;
    
    isReg   = range.getCell(cellRow, isRegCol).getValue();
    glLevel = range.getCell(cellRow, glLevelCol).getValue();
    glName  = range.getCell(cellRow, glNameCol).getValue();
    glFinalPoint  = range.getCell(cellRow, glFinalPointCol).getValue();

    if(isReg == "x" && glLevel == 3 && glName != "" && glFinalPoint >= 65.0) {
      Logger.log(id);
      range.getCell(cellRow, commDateCol).setValue(commDate);
      range.getCell(cellRow, commLocationCol).setValue(commLocation);
    }
    
    // update rowStart cell
    rowStartCell.setValue(cellRow);
  }
  
};



/////////////////////////////////////////////////////////////////////////////////////////////////////
//
// 04 - Update Confirmation date and location
//
/////////////////////////////////////////////////////////////////////////////////////////////////////
function updateConfirmationInfo() {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  
  var response = ui.alert(
      'Warning!!!',
      'Do you want to update Confirmation information for students?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response == ui.Button.YES) {
    updateConfirmationInfoImpl();
  }
}

function updateConfirmationInfoImpl() {
  var idCol       = 1;
  var glLevelCol  = 7;
  var glNameCol   = 8;
  var isRegCol    = 11;
  
  var confDateCol = 28;
  var confLocationCol= 29;
  var glFinalPointCol   = 32;  
  
  //////////////////////////// Get data from the Calendar sheet  //////////////////////////////
  var varSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Calendar");
  var varRange = varSheet.getRange(1, 1, 10, 11); //row, col, numRows, numCols 
  var confDate = varRange.getCell(4, 2).getValue(); //row, col
  var confLocation = "St. Maria Goretti, San Jose, CA"
  ////////////////////////////////////////////////////////////////////////////////////
  
  
  ////////////////////////////////////////////////////////////////////////////////////
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Std_VGz6v3");
  var range = sheet.getRange(1, 1, 700, 34); //row, col, numRows, numCols 
  var rowStartCell = range.getCell(2, 34);
  var rowStart = rowStartCell.getValue()+1;
  Logger.log(confDate);
  var id, isReg, glLevel, glName, glFinalPoint; 
  // iterate through all cells in the range
  for (var cellRow = rowStart; ; cellRow++) {
    id = range.getCell(cellRow, idCol).getValue(); 

    if(id == "") break;
    
    isReg   = range.getCell(cellRow, isRegCol).getValue();
    glLevel = range.getCell(cellRow, glLevelCol).getValue();
    glName  = range.getCell(cellRow, glNameCol).getValue();
    glFinalPoint  = range.getCell(cellRow, glFinalPointCol).getValue();

    if(isReg == "x" && glLevel == 8 && glName != "" && glFinalPoint >= 65) {
      range.getCell(cellRow, confDateCol).setValue(confDate);
      range.getCell(cellRow, confLocationCol).setValue(confLocation);
    }
    
    // update rowStart cell
    rowStartCell.setValue(cellRow);
  }
  
};





/////////////////////////////////////////////////////////////////////////////////////////////////////
//
// 05 - Un-share classes to the teachers
//
/////////////////////////////////////////////////////////////////////////////////////////////////////
function unShareClasses() {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  
  var response = ui.alert(
      'Warning!!!',
      'Do you want to stop sharing classes from the teachers?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response == ui.Button.YES) {
    shareClassesImpl(false);
  }
}



/////////////////////////////////////////////////////////////////////////////////////////////////////
//
// 06 - saveFormerStudents
//
/////////////////////////////////////////////////////////////////////////////////////////////////////
function saveFormerStudents() {
  
  var formerStudentFolderId = getStr("FORMER_STUDENT_FOLDER_ID");
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var file = DriveApp.getFileById(ss.getId());
  var clsFolder = DriveApp.getFolderById(formerStudentFolderId);
  
  // Make a copy
  var newFile = file.makeCopy("new file", clsFolder);

  // Open the new spreadsheet
  var ss = SpreadsheetApp.openById(newFile.getId());
  
  // Remove links to all class spreadsheets
  ss.getSheetByName("GL1A").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("GL1B").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("GL2A").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("GL2B").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("GL3A").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("GL3B").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("GL4A").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("GL4B").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("GL5A").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("GL5B").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("GL6A").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("GL6B").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("GL7A").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("GL7B").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("GL8A").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("GL8B").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("VN1A").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("VN1B").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("VN2A").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("VN2B").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("VN3A").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("VN3B").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("VN4A").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("VN4B").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("VN5A").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("VN5B").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("VN6A").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("VN6B").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("VN7A").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("VN7B").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("VN8A").getRange("O2:O2").getCell(1, 1).setValue("");
  ss.getSheetByName("VN8B").getRange("O2:O2").getCell(1, 1).setValue("");
  //ss.getSheetByName("gl-classes").getRange("G2:H17").clear();
  //ss.getSheetByName("vn-classes").getRange("G2:H17").clear();
  
  // Update First Communion sheet
  var temp = "=query(students!1:686, \"select A,B,C,D,E,F,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,AA,AB,AC,AD where G=3 and AD>=" + getStr("PASSING_POINT") + " order by C,E\")";
  ss.getSheetByName("Eucharist").getRange("A1:A1").getCell(1, 1).setValue(temp);

  // Update Confirmation sheet
  temp = "=query(students!1:686, \"select A,B,C,D,E,F,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,AA,AB,AC,AD where G=8 and AD>=" + getStr("PASSING_POINT") + " order by C,E\")";
  ss.getSheetByName("Confirmation").getRange("A1:A1").getCell(1, 1).setValue(temp);
  
  ss.getSheetByName("gl-honor-roll").clear();
  ss.getSheetByName("vn-honor-roll").clear();
};



/////////////////////////////////////////////////////////////////////////////////////////////////////
//
// 07 - updateClassesForNewReg
//
/////////////////////////////////////////////////////////////////////////////////////////////////////
function updateClassesForNewReg() {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  
  var response = ui.alert(
      'Warning!!!',
      'Do you want to set up classes for new registration process?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response == ui.Button.YES) {
    updateClassesForNewRegImpl();
  }
}

function waitSeconds(iMilliSeconds) {
    var counter= 0
        , start = new Date().getTime()
        , end = 0;
    while (counter < iMilliSeconds) {
        end = new Date().getTime();
        counter = end - start;
    }
}

function updateClassesForNewRegImpl() {
  
  var idCol       = 1;  
  var glGCol      = 7; 
  var glNCol      = 8; 
  var vnGCol      = 9; 
  var vnNCol      = 10; 
  var isRegCol    = 11; 
  
  var glFinalPointCol   = 32;  
  var vnFinalPointCol   = 33; 

  var passing_point = parseInt(getStr("PASSING_POINT"));
 
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Std_VGz6v3");
  var range = sheet.getRange(1, 1, 700, 34); //row, col, numRows, numCols
  var rowStartCell = sheet.getRange("AH1:AH1").getCell(1, 1); // <= Need to update column
  ////////////////////////////////////////////////////////////////////////////////////
 
  
  var id, isReg, glG, glN, vnG, vnN, glFinalPoint, vnFinalPoint; 
  // iterate through all cells in the range
  for (var cellRow = rowStartCell.getValue(); ; cellRow++) {
    id = range.getCell(cellRow, idCol).getValue(); 

    if(id == "") break;
    
    isReg   = range.getCell(cellRow, isRegCol).getValue();
    glG     = range.getCell(cellRow, glGCol).getValue();
    glN     = range.getCell(cellRow, glNCol).getValue();
    vnG     = range.getCell(cellRow, vnGCol).getValue();
    vnN     = range.getCell(cellRow, vnNCol).getValue();
    glFinalPoint  = range.getCell(cellRow, glFinalPointCol).getValue();
    vnFinalPoint  = range.getCell(cellRow, vnFinalPointCol).getValue();
    
    if(isReg == "x") {
      if(glG > 0 && glN != "" && glFinalPoint >= passing_point) {
        range.getCell(cellRow, glGCol).setValue(glG + 1);
      }
      
      if(vnG > 0 && vnN != "" && vnFinalPoint >= passing_point) {
        range.getCell(cellRow, vnGCol).setValue(vnG + 1);
      }
      
      // Clear the x
      range.getCell(cellRow, isRegCol).setValue('');
      
      // Clear glFinalPoint and vnFinalPoint
      
      waitSeconds(3000);
    }
   
    // update rowStart cell
    rowStartCell.setValue(cellRow);
    
    //Logger.log(cellRow);
  }
  
};


