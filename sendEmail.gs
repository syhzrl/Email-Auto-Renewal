//global var
var app = SpreadsheetApp.openById('1Np3UhwCmN9hqWHGrhrFOIxteeFaDkU55q0lWurSE_EQ');
var uploadSheet = app.getSheetByName('Upload');
var templateSheet = app.getSheetByName('Email Template');
var reconSheet = app.getSheetByName('Recon');
var reconLastRow = reconSheet.getRange('Recon!A1').getDataRegion().getLastRow();


function sendEmail() {
  
  for (let i = 2; i < reconLastRow+1; i++){
    
    // dynamic member's info
    let dynName = reconSheet.getRange(i,1).getValue();             //member's Name
    let dynCompany = reconSheet.getRange(i,2).getValue();          //member's Company
    let dynEndDate = reconSheet.getRange(i,4).getDisplayValue();   //member's End Date
    let locationChecker = reconSheet.getRange(i,6).getValue();     //member's Location
    let status = reconSheet.getRange(i,7).getDisplayValue();       //member's Expiry Status
    let statusChecker = status.includes('Not Expiring');           //boolean of status
    
    //email info to send  
    let emailTo = reconSheet.getRange(i,3).getValue();              //member's email
    let emailSubject = 'Membership Renewal ('+dynCompany+')';       //email subject
    let htmlBody = HtmlService.createHtmlOutputFromFile('emailBody').getContent();
    let emailBody =                                                 //email body link to emailBody.html
        htmlBody
        .replace('{Name}',dynName)                                                
        .replace('{End Date}',dynEndDate)
        .replace('{Location}',locationChecker);
    
    if(!statusChecker) {                                           //condition to only send email to expiring members
      if(locationChecker === 'WORQ Subang') {
        MailApp.sendEmail(emailTo,emailSubject,'body',{htmlBody:emailBody,replyTo:templateSheet.getRange(2,2).getValue(),name:locationChecker,cc:templateSheet.getRange(4,2).getValue()})
 
      }

      if(locationChecker === 'WORQ TTDI') {
        MailApp.sendEmail(emailTo,emailSubject,'body',{htmlBody:emailBody,replyTo:templateSheet.getRange(7,2).getValue(),name:locationChecker,cc:templateSheet.getRange(9,2).getValue()})

      }

      if(locationChecker === 'WORQ Gateway') {
        MailApp.sendEmail(emailTo,emailSubject,'body',{htmlBody:emailBody,replyTo:templateSheet.getRange(12,2).getValue(),name:locationChecker,cc:templateSheet.getRange(14,2).getValue()})

      }

        
      if(locationChecker === 'WORQ Surian') {
        MailApp.sendEmail(emailTo,emailSubject,'body',{htmlBody:emailBody,replyTo:templateSheet.getRange(17,2).getValue(),name:locationChecker,cc:templateSheet.getRange(19,2).getValue()})

      }

      reconSheet.getRange(i,8).setValue('Sent');  
    }
  }
}




