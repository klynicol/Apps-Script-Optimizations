/*************************************************************************************************

    ---------Marination Mobile LLC, Load List Script---------
    --------------Created by Mark Wickline 2017-------------

Functionality:
    -Creates a new sheet from the template every Thursday morning @ 4:00AM and names it to the appropriate date range. (Template does not have to be visible to the user for this to happen).
    -Automatically hides columns and old sheets when document is opened.
    -Updates nightly @ 4:00 AM to expand all the columns and hide old sheets. Updates current sheet to hide columns from previous day.
    -Adds button functionality to expand/contract columns based on the day of the week.
    -Adds a "DONE" button that will send an email to inform recipients load list is complete and who completed it.
    
Additional updates (via built in google sheets functionality):
    -Added permissions for protection from editing "pars" and "to load" columns in gray. Permissions can be set per google account
    -Add MAX function to all the gray columns to avoid showing negative numbers.
    
***************************************************************************************************/


/*************************************************************
--------EDIT VALUES ------------------------------------------
-------- null=none--------------------------------------------
*/

var email_recipient_1 = "roz@marinationmobile.com";
var email_recipient_2 = "klynicol@gmail.com";
var email_recipient_3 = null;
  

/************************************************************/

/**
 * @OnlyCurrentDoc
 */

function onOpen(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lastSheet = ss.getSheets()[ss.getSheets().length - 1];
  if(lastSheet.isSheetHidden()) lastSheet.showSheet();
  ss.setActiveSheet(lastSheet);
  hideSheets();
  hideColumns();
}


function daySelect(){
 var d = new Date();
 return d.getDay(); //0 = Sunday
}


function hideSheets(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allSheets = ss.getSheets();
  for(index = 0; index < allSheets.length - 1; index++){
    if(!allSheets[index].isSheetHidden()){
      allSheets[index].hideSheet();
    }
  }
}


function cloneTemplate() {                                        //Runs every Thursday morning @ 4:00 AM
  var d = new Date();
  var month = d.getMonth() + 1;
  var day = d.getDate();
  var calculateDate = d.getTime() + 518400000;                    //adding 6 days
  var dEnd = new Date();
  dEnd.setTime(calculateDate);
  var monthEnd = dEnd.getMonth() + 1;
  var dayEnd = dEnd.getDate();
  var name = month + "." + day + " - " + monthEnd + "." + dayEnd; //Marination Current format
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Template').copyTo(ss);
  SpreadsheetApp.flush();
  sheet.setName(name);
  ss.setActiveSheet(sheet);
  
  //copy the cell protection over from the template
  var rangeTop = ss.getRange('E1:Y4');
  rangeTop.protect().setDescription('Top').setWarningOnly(true);
  var rangeLeft = ss.getRange('A1:D66');
  rangeLeft.protect().setDescription('Left').setWarningOnly(true);
  var rangeTotals = ss.getRange('W5:Y40');
  rangeTotals.protect().setDescription('Totals').setWarningOnly(true);
  var rangeCatering = ss.getRange('E42:V43');
  rangeCatering.protect().setDescription('Catering Days').setWarningOnly(true);
  //gray columns protection
  var rangeSunday = ss.getRange('L5:L40');
  rangeSunday.protect().setDescription('Sunday').setWarningOnly(true);
  var rangeMonday = ss.getRange('O5:O40');
  rangeMonday.protect().setDescription('Monday').setWarningOnly(true);
  var rangeTuesday = ss.getRange('F5:F40');
  rangeTuesday.protect().setDescription('Tuesday').setWarningOnly(true);
  var rangeWednesday = ss.getRange('U5:U40');
  rangeWednesday.protect().setDescription('Wednesday').setWarningOnly(true);
  var rangeThursday = ss.getRange('F5:F40');
  rangeThursday.protect().setDescription('Thursday').setWarningOnly(true);
  var rangeFriday = ss.getRange('I5:I40');
  rangeFriday.protect().setDescription('Friday').setWarningOnly(true);

  if(sheet.isSheetHidden()){
    sheet.showSheet();
  }
  hideSheets();
  showColumns();
  hideColumns();
}



function sendEmail(){
  
  /**************************************************************************
  EDIT THESE FOR RECIPIENTS
  ***************************************************************************/
  
  var ui = SpreadsheetApp.getUi();
  var user = Session.getActiveUser().getEmail();
  MailApp.sendEmail(email_recipient_1, "Amazon Load List Complete", "Completed by: " + user);
  MailApp.sendEmail(email_recipient_2, "Amazon Load List Complete", "Completed by: " + user);
  MailApp.sendEmail(email_recipient_3, "Amazon Load List Complete", "Completed by: " + user);
  ui.alert("Email sent to: " + email_recipient_1);
}

//omited from Amazon Load List Script because it's causing permission issues.


function nightlyUpdate(){                                       //Runs every night @ 4:00 AM
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allSheets = ss.getSheets();
  var lastSheet = ss.getSheets()[ss.getSheets().length - 1];
  if(lastSheet.isSheetHidden()) lastSheet.showSheet();          //unhide current sheet
  lastSheet.showColumns(1,25);
  for(index = 0; index < allSheets.length - 1; index++){
    if(!allSheets[index].isSheetHidden()){
      allSheets[index].hideSheet();
    }
    allSheets[index].showColumns(1,25);                         //run expand on all old sheets
  }
  hideColumnsSwitch(daySelect(), lastSheet);
}



function showColumns(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  sheet.showColumns(1,25);
  Logger.log(sheet);
}


function hideColumns(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  if(sheet){
    var d = daySelect();
    hideColumnsSwitch(d, sheet);
    Logger.log(sheet);
  }
  else Logger.log("Error, hideColumns did not find active sheet");
}


function hideColumnsSwitch(day, sheet){
  switch (day){
      case 0: //Sunday
        sheet.hideColumns(1);
        sheet.hideColumns(5,6);
        sheet.hideColumns(14,12);
        break;
      case 1: //Monday
        sheet.hideColumns(1);
        sheet.hideColumns(5,9);
        sheet.hideColumns(17,9);
        break;
      case 2: //Tuesday
        sheet.hideColumns(1);
        sheet.hideColumns(5,12);
        sheet.hideColumns(20,6);
        break;
      case 3: //Wednesday
        sheet.hideColumns(1);
        sheet.hideColumns(5,15);
        sheet.hideColumns(23,3);
        break;
      case 4: //Thursday
        sheet.hideColumns(1);
        sheet.hideColumns(8,18);
        break;
      case 5: //Friday
        sheet.hideColumns(1);
        sheet.hideColumns(5,3);
        sheet.hideColumns(11,15);
        break;
      default: // default and saturday
        sheet.hideColumns(1);
        sheet.hideColumns(23,3);
        break;
    }
}

