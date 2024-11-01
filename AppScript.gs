// Global variables
  var  consultationNum;    
  var  date; 
  var  venue;
  var  timeStart;  
  var  timeEnd;      
  var  details;      
  
  var templateID = "1fTdKOD6RUCKDCf4r5AqkaMUUxUkWzJHAYkV2zsI2FI8";
  //https://docs.google.com/document/d/1fTdKOD6RUCKDCf4r5AqkaMUUxUkWzJHAYkV2zsI2FI8/edit?usp=sharing
  var folderID = "1t-_SrhOOEguFqk59CQELaeGIY9sdWTvl"
  //https://drive.google.com/drive/folders/1t-_SrhOOEguFqk59CQELaeGIY9sdWTvl?usp=sharing
 function processDocuments(){

  //response sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName('Form Responses 1');

  dataSheet.sort(1, false);
  Utilities.sleep(3000);



  var calendar = CalendarApp.getDefaultCalendar();
  var timeZone=  calendar.getTimeZone();
  consultationNum = dataSheet.getRange('B2').getValue();    
  date     = dataSheet.getRange('C2').getValue(); 
  date     = Utilities.formatDate(date, "GMT+08", 'MMMMMMMMM dd, yyyy');
  venue    = dataSheet.getRange('D2').getValue();
  timeStart= dataSheet.getRange('E2').getValue(); 
  timeStart=Utilities.formatDate(timeStart, "GMT+08","hh:mm aaa" );  
  timeEnd  = dataSheet.getRange('F2').getValue();     
  timeEnd  = Utilities.formatDate(timeEnd,"GMT+08", "hh:mm aaa" );
  details  = dataSheet.getRange('G2').getValue(); 

  var originalDoc = DriveApp.getFileById(templateID);
  var folder = DriveApp.getFolderById(folderID);

  var formBlob = originalDoc.makeCopy('Form'+consultationNum, folder);
  var formId = formBlob.getId();

  var formDoc = DocumentApp.openById(formId);
  var body = formDoc.getBody();


  var replacements = {
    '<consultationNum' : consultationNum,    
    '<date' : date, 
    '<venue': venue,
    '<timeStart': timeStart,  
    '<timeEnd': timeEnd,      
    '<details': details      
   };
    
  // Loop through the replacements and perform the find and replace
  for (var oldText in replacements) {
    var newText = replacements[oldText];
    body.replaceText(oldText, newText);
  }
  formDoc.saveAndClose(); 
  
  const blob = formDoc.getAs('application/pdf');
  const file = folder.createFile(blob);

  }
  //************************************************************************************* */
  
  


