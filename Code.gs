/*
    TO-DO:
      - Confirm CREATED writing at completion of Task
      - Finish Documenting Code
*/

function create_exams(){
  
  var ss = SpreadsheetApp.getActiveSheet(); //Gets the main spreadsheet for reference
  
  // -------------------------------------- KEEP THIS SECTION UP TO DATE ---------------------------------------------------------
 /********************************************************************************************************************************
                                                                                                                                
   - This section declares all the ranges (locations) of data on the physical Google Sheet that are used in this script        
   - This section should be updated when you move the location of items/ add rows or columns etc aka physical changes to sheet 
                                                                                                                                
  ********************************************************************************************************************************/
  
  var files_range = "Settings!I3:I16";
  var extras_range = "Settings!I19:I30";
  var generator_range = "A5:O";

  var files = ss.getRange(files_range).getValues();
  var master_list = files[0]; //The position on the settings page containing the master list sheetID
  var icq_list = files[1]; //The position on the settings page containing the master list sheetID 

  var extras = ss.getRange(extras_range).getValues();
  var semester = extras[0];
  var signature = extras[1];
  var tableStartRow = extras[2]; //The number of the first row for input values on the main table
  var createdColumn = extras[3];
  var examineeEmailColumn = extras[4]; //The column number for examinee's email sent value
  var examinerEmailColumn = extras[5]; //The column number for the examiner's email sent value


/********************************************************************************************************************************* 
********************************************************************************************************************************* */ 

  

  

  //Get the Master Exam List spreadsheet from the value (spreadsheetID) in 'master_list_range'
  var examMaster = SpreadsheetApp.openById(master_list); 

  //Get the ICQ Exam List spreadsheet from the value (spreadsheetID) in 'icq_list_range'
  var icqMaster = SpreadsheetApp.openById(icq_list);
  

  var data = ss.getRange(generator_range).getValues();//Read all the data from generator range
  
  //Loops through each row in data, generating the exam and sending the emails
  for (var i = 0; i < data.length ; i++){
    var row = parseDataRow(data[i]);
    if(validRow(row)){
      createExamSpreadsheet(row,examMaster,icqMaster,ss,i+tableStartRow,semester, createdColumn);
      sendEmails(row,ss,i+tableStartRow,signature,examineeEmailColumn,examinerEmailColumn);
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
    - semester: The semester to record
*/
function createExamSpreadsheet(row,examMaster,icqMaster,ss,i,semester, createdColumn){

  var folder= null;
  
  if(row.created!="DONE"){ //Don't recreate exam if its already marked as done
    //Logger.log(row.examFolder);
    //Create exam folder if it doesn't already exist
    if(row.examFolder==""){
      var newFolder = DriveApp.getFolderById(row.pr).createFolder(row.examineeName+" (EXAMS)"); //Create a new folder within the SiTs PR folder named NAME (EXAMS)
      var id = newFolder.getId(); //Get the Drive FolderId
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SiT Information").getRange(row.inforow, 5).setValue(id); //Store the folder ID on the Info sheet
      folder = newFolder;
    }
    else{
      folder = DriveApp.getFolderById(row.examFolder); //If a folder is already saved on the SiT info sheet, use that one
    }
    
    var template = DriveApp.getFileById(row.examId); //Get the template's file from Drive SheetId

    var date = Utilities.formatDate(new Date(row.date), 'America/New_York', 'MM/dd/yyyy')
    var examRubric = template.makeCopy(row.examineeNum+" "+row.exam+" "+date+" ["+row.examinerNum+"]",folder); //make a copy of the template to the SiTs exam folder

    examRubric.addEditor(row.examinerEmail); //Add edit accesss for examiner
    examRubric.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); //Allow View-only access with link sharing
    
    var masterSheet = null;

    if(row.exam == "ICQ"){ //Use the ICQ master sheet if ICQ exam
      masterSheet = icqMaster.getSheetByName("All Exams");
    }
    else{ //Else use the regular sheet
      masterSheet = examMaster.getSheetByName("All Exams");
    }

    writeToMasterList(row,masterSheet,semester,examRubric.getUrl());
    
    ss.getRange(i,createdColumn).setValue("DONE"); //Mark exam generation as done
  }

}

/*
  Function responsible for writing to necessary information to the exam list spreadsheet
    - row: the data object for the current exam row being created
    - masterSheet: The sheet that the exam links and info get written to
    - Semester: The semester to write the exam as
  */
function writeToMasterList(row,masterSheet,semester,url){

  //Semester, date, exam, examinee, examiner, link
  var newRow = masterSheet.getLastRow()+1; //Write to the bottom of the list
  var newRange = masterSheet.getRange(newRow,1,1,6);
  newRange.setValues([[semester,row.date,row.exam,row.examineeName+" ("+row.examineeNum+")",row.examinerName+" ("+row.examinerNum+")",url]]);
}

/*
  Function responsible for sending the emails to the examiner/examinee
    - row: the data object for the current exam row being created
    - ss: The current spreadsheet
    - i: the row on the sheet being worked on
*/
function sendEmails(row,ss,i,signature,examineeEmailColumn, examinerEmailColumn){

  var subject = "RSP "+row.exam+" Exam"; // The subject Line
  var date = Utilities.formatDate(new Date(row.date), 'America/New_York', 'MM/dd/yyyy');

  //Message for Examinee
  var examineeEmail = `<p>Your <b>${row.exam} Exam</b>  has been scheduled for <b>${date}</b> with <b>${row.examinerNum}</b>.</p>`+
                `<p>\n\nIf you have any questions and/or concerns, please reach out to me</p>`+
                `<p>\n\nThank you,<br\>\n${signature}</p>`;

  //Message for examiner
  var examinerEmail = `<p>A <b>${row.exam} Exam</b>  has been scheduled for <b>${date}</b> with <b>${row.examineeNum}</b>.</p>`+
                    `<p>\n\nThe exam rubric has been shared with you via Google Drive.</p>`+
                    `<p>\n\nIf you have any questions and/or converns, please reach out to me.</p>`+
                    `<p>\n\nThank you,\n<br/>${signature}</p>`;

  var examineePlain = examineeEmail.replace(/(<([^>]+)>)/ig, ""); // clear html tags for plain mail
  

  if(row.examineeEmailSent!="SENT"){ //If not marked as completed. send examinee email and mark as done
    MailApp.sendEmail(row.examineeEmail, subject,examineePlain,{ htmlBody: examineeEmail });
    ss.getRange(i,examineeEmailColumn).setValue("SENT");
  }

  var examinerPlain = examinerEmail.replace(/(<([^>]+)>)/ig, ""); // clear html tags for plain mail

  if(row.examinerEmailSent!="SENT"){ //If not marked as completed, send examiner email and mark as done
    MailApp.sendEmail(row.examinerEmail,subject,examinerPlain,{ htmlBody: examinerEmail });
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
