function create_exams(){
  var ss = SpreadsheetApp.getActiveSheet();
  
  var dc = "Miles Foydl (120)"
  
  //Gets the file associated with the exam template, to change these, change the ID's at the top of the spreadsheet
  
  var pnp = DriveApp.getFileById(ss.getRange("C1").getValue()); //Template File for P&P Exam
  var sup = DriveApp.getFileById(ss.getRange("C2").getValue()); //Template File for SUP Exam
  var dis = DriveApp.getFileById(ss.getRange("C3").getValue()); //Template File for DIS Exam
  var icq = DriveApp.getFileById(ss.getRange("C4").getValue());

  var master = SpreadsheetApp.openById(ss.getRange("C5").getValue()); //Get the current master exam list
  var icqMaster = SpreadsheetApp.openById(ss.getRange("C6").getValue());
  
  var root_folder = DriveApp.getFolderById(ss.getRange("C7").getValue()); //Root Folder for current SiT Exam Folders
  
  
  var startRow = 9;
  var numRows= ss.getLastRow()-startRow;
  var numCols= 11;
  var range = ss.getRange("A9:N28"); //The feasible range that values could be placed
  var data = range.getValues()
  var examNames=["P&P","SUP","DIS","ICQ"]; //These are the names of the exams given, should match values in spreadsheet exactly
  
  for (var i = 0; i < numRows ; i++){
    Logger.log(i);
    //var exam = range.getCell(i,1).getValue();
    var exam = data[i][0];
    var date = Utilities.formatDate(new Date(data[i][1]), 'America/New_York', 'MM/dd/yyyy');
    var examineeNum= data[i][2];
    var examineeName = data[i][3];
    var examineeEmail = data[i][4];
    var examinerNum = data[i][5];
    var examinerName = data[i][6];
    var examinerEmail = data[i][7];
    var sent1 = range.getCell(i+1,9);
    var sent2 = range.getCell(i+1,10);
    var created = range.getCell(i+1,11);
    var pr = data[i][11];
    var folder = data[i][12];
    var row= data[i][13];
    
    
    
   
   
    
    if(exam!=""&&examineeNum!=""&&examinerNum!=""){
      if(created.getValue()!="DONE"){
        //Create exam folder if it doesn't already exist
        if(folder==""){
          var newFolder = DriveApp.getFolderById(pr).createFolder(examineeName+" (EXAMS)");
          var id = newFolder.getId();
          SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SiT Information").getRange(row, 5).setValue(id);
          folder = newFolder;
          data[i][12]=id;
        }
        else{
          folder = DriveApp.getFolderById(folder);
        }
        
        
        var examRubric = null;
        switch(exam){ //This section creates a copy of the exam template and names it "S# Exam mm/dd/yyyy [DC/SL#]"
          case examNames[0]:
            examRubric = pnp.makeCopy(examineeNum + " " + exam + " " + date + " ["+examinerNum+"]",folder);
            break;
          case examNames[1]:
            examRubric = sup.makeCopy(examineeNum + " " + exam + " " + date + " ["+examinerNum+"]",folder);
            break;
          case examNames[2]:
            examRubric = dis.makeCopy(examineeNum + " " + exam + " " + date + " ["+examinerNum+"]",folder);
            break;
          case examNames[3]:
            examRubric = icq.makeCopy(examineeNum + " " + exam + " "+ date + " ["+examinerNum+"]",folder);
            break;
            
        }
        examRubric.addEditor(examinerEmail); //Add edit accesss for examiner
        examRubric.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); //Allow link sharing
        
        var masterSheet = master.getSheetByName("All Exams");
        if (exam == examNames[3]){
          masterSheet = icqMaster.getSheetByName("All Exams");
        }
        var newRow = masterSheet.getLastRow()+1;
        var newRange = masterSheet.getRange(newRow,1,1,5);
        newRange.getCell(1,1).setValue(date);
        newRange.getCell(1,2).setValue(exam);
        newRange.getCell(1,3).setValue(examineeName+" ("+examineeNum+")");
        newRange.getCell(1,4).setValue(examinerName+" ("+examinerNum+")");
        newRange.getCell(1,5).setValue(examRubric.getUrl());
      }
      created.setValue("DONE");
      
      //Message for examinee
      var message = "<p>Your <b>"+exam+"</b> exam has been scheduled for <b>"+date+"</b> with "+examinerNum+".</p>"+ "<p>\n\nIf you have any questions and/or concerns, please reach out to me</p> <p>\n\nThank You,<br/>\n"+dc+"</p>";
      var msgPlain = message.replace(/(<([^>]+)>)/ig, ""); // clear html tags for plain mail
      var subject = "RSP "+exam+" Exam";
      if(sent1.getValue()!="SENT"){
        MailApp.sendEmail(examineeEmail, subject,msgPlain,{ htmlBody: message });
        sent1.setValue("SENT");
      }
      //Message for examiner
      var message2 = "<p>A <b>"+exam+"</b> exam has been scheduled for <b>"+date+"</b> with "+examineeNum+".</p>"+
                     "<p>\n\nThe exam rubric has been shared with you via google drive.</p>"+
                     "<p>\n\nIf you have any questions and/or concerns, please reach out to me.</p>"+
                       "<p>\n\nThank You,\n<br/>"+dc+"</p>"+"";
                     //'<p><img src="https://media.giphy.com/media/pt0EKLDJmVvlS/giphy.gif"></p>';
      var msg2Plain = message2.replace(/(<([^>]+)>)/ig, ""); // clear html tags for plain mail
      
      if(sent2.getValue()!="SENT"){
        MailApp.sendEmail(examinerEmail,subject,msg2Plain,{ htmlBody: message2 });
        sent2.setValue("SENT");
      }
      
      
      //Where to store all exams? add this to copy_file line for folder ID
    }
    
  }
}
