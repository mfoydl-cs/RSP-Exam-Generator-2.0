/*
    TO-DO:
      - Confirm CREATED writing at completion of Task
      - See if there's a better way to pass variables/abstract
      - Try to remove any uncessary calls to sheets API
      - Finish Documenting Code

      - See if email/create could use any better reworks
      - Rework the write to master list function!
      
      Future Ideas:
         - write an archive function to keep the lists clean
         - sepparate by semester?!

*/

function create_exams(){

  
  // -------------------------------------- KEEP THIS SECTION UP TO DATE ---------------------------------------------------------
 /********************************************************************************************************************************
                                                                                                                                
   - This section declares all the ranges (locations) of data on the physical Google Sheet that are used in this script        
   - This section should be updated when you move the location of items/ add rows or columns etc aka physical changes to sheet 
                                                                                                                                
  ********************************************************************************************************************************/
  
  //Redo this to be similar to getting rows -- do for all the settings - less get ranges

  var master_list_range = "Settings!I3"; //The position on the settings page containing the master list sheetID
  var icq_list_range = "Settings!I4"; //The position on the settings page containing the master list sheetID 
  var generator_range = "A5:O";

  var tableStartRow = 5; //The number of the first row for input values on the main table

/********************************************************************************************************************************* 
********************************************************************************************************************************* */ 

  

  var ss = SpreadsheetApp.getActiveSheet(); //Gets the main spreadsheet for reference

  //Get the Master Exam List spreadsheet from the value (spreadsheetID) in 'master_list_range'
  var examMaster = SpreadsheetApp.openById(ss.getRange(master_list_range).getValue()); 

  //Get the ICQ Exam List spreadsheet from the value (spreadsheetID) in 'icq_list_range'
  var icqMaster = SpreadsheetApp.openById(ss.getRange(icq_list_range).getValue());
  

  var data = ss.getRange(generator_range).getValues();//Read all the data from generator range
  
  //Loops through each row in data, generating the exam and sending the emails
  for (var i = 0; i < data.length ; i++){
    var row = parseDataRow(data[i]);
    if(validRow(row)){
      createExamSpreadsheet(row,examMaster,icqMaster,ss,i+tableStartRow);
      sendEmails(row,ss,i+tableStartRow);
    }
    
  }
}



/*
  This function turns a row of data from the 'Generator' Page of the spreadsheet, into an object for the rest of the script to use
  This will make it so that if the columns are changed on the spreadsheet, the only thing that needs to be changed is the columns array
*/
function parseDataRow(row){
  // -------------------------------------- KEEP THIS SECTION UP TO DATE ---------------------------------------------------------
 /********************************************************************************************************************************
                                                                                                                                
   - This section lists the orders of all values in the spreadsheet!! The values listed are the variable names used in the 
      program to represent each column in the data table.        
                                                                                                                                
  ********************************************************************************************************************************/
  //Please have this array replicate the exact order of every column in the 'Generator' main sheet!!!!!
  var columns = ["exam","date","examineeNum","examineeName","examineeEmail",
                  "examinerNum","examinerName","examinerEmail","examineeEmailSent",
                  "examinerEmailSent","created","pr","examFolder","inforow","examId"];

  /********************************************************************************************************************************* 
**********************************************************************************************************************************/

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
    - i: the row number on the spreadsheet
*/
function createExamSpreadsheet(row,examMaster,icqMaster,ss,i){
  var folder= null;
  var createdColumn = 11;
  if(row.created!="DONE"){
    //Create exam folder if it doesn't already exist
    if(row.folder==""){
      var newFolder = DriveApp.getFolderById(row.pr).createFolder(row.examineeName+" (EXAMS)");
      var id = newFolder.getId();
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SiT Information").getRange(row.inforow, 5).setValue(id);
      folder = newFolder;
    }
    else{
      folder = DriveApp.getFolderById(row.folder);
    }
    
    var template = DriveApp.getFileById(row.examId);
    var examRubric = template.makeCopy(row.examineeNum+" "+row.exam+" "+row.date+" ["+row.examinerNum+"]",folder);

    examRubric.addEditor(row.examinerEmail); //Add edit accesss for examiner
    examRubric.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); //Allow View-only access with link sharing
    
    var masterSheet = Null;

    if(row.exam == "ICQ"){
      masterSheet = icqMaster.getSheetByName("All Exams");
    }
    else{
      masterSheet = examMaster.getSheetByName("All Exams");
    }

    writeToMasterList(row,masterSheet);
    
    ss.getRange(i,createdColumn).setValue("DONE");
  }

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
function sendEmails(row,ss,i){
  var dc = "Miles Foydl (120)"
  var examinerEmailColumn = 10;
  var examineeEmailColumn = 9;
  //Message for examinee

  var message = `<p>Your <b>${row.exam} Exam</b>  has been scheduled for <b>${row.date}</b> with <b>${row.examinerNum}</b>.</p>`+
                `<p>\n\nIf you have any questions and/or concerns, please reach out to me</p>`+
                `<p>\n\nThank you,<br\>\n${dc}</p>`;

  var msgPlain = message.replace(/(<([^>]+)>)/ig, ""); // clear html tags for plain mail
  var subject = "RSP "+row.exam+" Exam";

  if(row.examineeEmailSent!="SENT"){
    MailApp.sendEmail(row.examineeEmail, subject,msgPlain,{ htmlBody: message });
    ss.getRange(i,examineeEmailColumn).setValue("SENT");
  }
  //Message for examiner
  var message2 = `<p>A <b>${row.exam} Exam</b>  has been scheduled for <b>${row.date}</b> with <b>${row.examineeNum}</b>.</p>`+
                    `<p>\n\nThe exam rubric has been shared with you via Google Drive.</p>`+
                    `<p>\n\nIf you have any questions and/or converns, please reach out to me.</p>`+
                    `<p>\n\nThank you,\n<br/>${dc}</p>`;
                  //'<p><img src="https://media.giphy.com/media/pt0EKLDJmVvlS/giphy.gif"></p>';

  var msg2Plain = message2.replace(/(<([^>]+)>)/ig, ""); // clear html tags for plain mail

  if(row.examinerEmailSent!="SENT"){
    MailApp.sendEmail(row.examinerEmail,subject,msg2Plain,{ htmlBody: message2 });
    ss.getRange(i,examinerEmailColumn).setValue("SENT");
  }
}

/*
  Function is responsible for checking that a data row is valid before processing it
    - row: the data object for the current exam row being created
*/
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
  var row = ["P&P","1/28/2021","S4","xxx","xxx","104","xxx","xxx","","","","Test","test2","18","nalskdjflskjdf"];
  var row2 = parseDataRow(row);
  Logger.log(row2);
}