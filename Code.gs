function create_exams(){
  var ss = SpreadsheetApp.getActiveSheet();
  
 /********************************************************************************************************************************
                                                                                                                                
   - This section declares all the ranges (locations) of data on the physical Google Sheet that are used in this script        
   - This section should be updated when you move the location of items/ add rows or columns etc aka physical changes to sheet 
                                                                                                                                
  ********************************************************************************************************************************/
  
  var master_list_range = "Settings!G3";
  var icq_list_range = "Settings!G4";

  var generator_range = "A5:O";
  
  //Get the Master Exam List spreadsheet from the value (spreadsheetID) in 'master_list_range'
  var examMaster = SpreadsheetApp.openById(ss.getRange(master_list_range).getValue()); 

  //Get the ICQ Exam List spreadsheet from the value (spreadsheetID) in 'icq_list_range'
  var icqMaster = SpreadsheetApp.openById(ss.getRange(icq_list_range).getValue());
  
  //var root_folder = DriveApp.getFolderById(ss.getRange("C7").getValue()); //Root Folder for current SiT Exam Folders
  
  
  //var startRow = 9;
  //var numRows= ss.getLastRow()-startRow;
  //var numCols= 11;
  //var range = ss.getRange("A9:N28"); //The feasible range that values could be placed

  var data = ss.getRange(generator_range).getValues();
  
  for (var i = 0; i < numRows ; i++){
    var row = parseDataRow(data[i]);
    if(validRow(row)){
      createExamSpreadsheet(row,examMaster,icqMaster);
      sendEmails(row);
    }
    
    }
}
/*
  This function turns a row of data from the 'Generator' Page of the spreadsheet, into an object for the rest of the script to use
  This will make it so that if the columns are changed on the spreadsheet, the only thing that needs to be changed is the columns array
*/
function parseDataRow(row){
  //Please have this array replicated the exact order of every column in the 'Generator' main sheet
  var columns = ["exam","date","examineeNum","examineeName","examineeEmail",
                  "examinerNum","examinerName","examinerEmail","examineeEmailSent",
                  "examinerEmailSent","created","pr","examFolder","inforow","examId"];

  var row_object = {};

  for(var i=0;i<columns.length;i++){
    row_object[columns[i]]=row[i];
  }

  return row_object;
}

/*
  Function responsible for making a copy of the Template Rubric into the examinee exam folder and sharing it
    - row: the data object for the current exam row being created
    - examMaster: the spreadsheet containing all the links to exams
    - icqMaster: the spreadsheet containing all the links to ICQ exams
*/
function createExamSpreadsheet(row,examMaster,icqMaster){
  var folder= null;
  if(row.created!="DONE"){
    //Create exam folder if it doesn't already exist
    if(row.folder==""){
      var newFolder = DriveApp.getFolderById(row.pr).createFolder(row.examineeName+" (EXAMS)");
      var id = newFolder.getId();
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SiT Information").getRange(row.inforow, 5).setValue(id);
      folder = newFolder;
      //data[i][12]=id;
    }
    else{
      folder = DriveApp.getFolderById(row.folder);
    }
    
    var template = DriveApp.getFileById(row.examId);
    var examRubric = template.makeCopy(row.examineeNum+" "+row.exam+" "+row.date+" ["+row.examinerNum+"]",folder);

    examRubric.addEditor(row.examinerEmail); //Add edit accesss for examiner
    examRubric.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); //Allow View-only access with link sharing
    
    var masterSheet = examMaster.getSheetByName("All Exams");

    if (row.exam == "ICQ"){
      masterSheet = icqMaster.getSheetByName("All Exams");
    }

    writeToMasterList(row,masterSheet);
    
  }
  //created.setValue("DONE");
}

/*
  Function responsible for writing to necessary information to the exam list spreadsheet
    - row: the data object for the current exam row being created
    - masterSheet: The sheet that the exam links and info get written to
  */
function writeToMasterList(row,masterSheet){
  var newRow = masterSheet.getLastRow()+1;
  var newRange = masterSheet.getRange(newRow,1,1,5);
  newRange.getCell(1,1).setValue(row.date);
  newRange.getCell(1,2).setValue(row.exam);
  newRange.getCell(1,3).setValue(row.examineeName+" ("+examineeNum+")");
  newRange.getCell(1,4).setValue(row.examinerName+" ("+examinerNum+")");
  newRange.getCell(1,5).setValue(row.examRubric.getUrl());
}

/*
  Function responsible for sending the emails to the examiner/examinee
    - row: the data object for the current exam row being created
*/
function sendEmails(row){
  var dc = "Miles Foydl (120)"
  //Message for examinee
      var message = "<p>Your <b>"+row.exam+"</b> exam has been scheduled for <b>"+row.date+"</b> with "+row.examinerNum+".</p>"+ "<p>\n\nIf you have any questions and/or concerns, please reach out to me</p> <p>\n\nThank You,<br/>\n"+dc+"</p>";
      var msgPlain = message.replace(/(<([^>]+)>)/ig, ""); // clear html tags for plain mail
      var subject = "RSP "+row.exam+" Exam";

      if(row.sent1!="SENT"){
        MailApp.sendEmail(row.examineeEmail, subject,msgPlain,{ htmlBody: message });
        //sent1.setValue("SENT");
      }
      //Message for examiner
      var message2 = "<p>A <b>"+row.exam+"</b> exam has been scheduled for <b>"+row.date+"</b> with "+row.examineeNum+".</p>"+
                     "<p>\n\nThe exam rubric has been shared with you via google drive.</p>"+
                     "<p>\n\nIf you have any questions and/or concerns, please reach out to me.</p>"+
                       "<p>\n\nThank You,\n<br/>"+dc+"</p>"+"";
                     //'<p><img src="https://media.giphy.com/media/pt0EKLDJmVvlS/giphy.gif"></p>';
      var msg2Plain = message2.replace(/(<([^>]+)>)/ig, ""); // clear html tags for plain mail
      
      if(row.sent2!="SENT"){
        MailApp.sendEmail(row.examinerEmail,subject,msg2Plain,{ htmlBody: message2 });
        //sent2.setValue("SENT");
      }
}
function validRow(row){
  if(row.exam==""){
    return false;
  }
  if(row.date==""){
    return false;
  }
  if(row.examineeNum==""){
    return false;
  }
  if(row.examinerNum==""){
    return false;
  }
  return true;
}
function test(){
  var row = ["P&P","1/28/2021","S4","Cheng, Joan","joan.cheng@stonybrook.edu","104","Zoa, Connie","connie.zao@@stonybrook.edu","","","","Test","test2","18","nalskdjflskjdf"];
  var row2 = parseDataRow(row);
  Logger.log(row2);
}