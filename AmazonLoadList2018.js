/*************************************************************************************************

    ---------Marination Mobile LLC, Load List Script---------
    --------------Created by Mark Wickline 2018--------------
    
**---Edit Parameters---***************************************************************************/

var COMMISSARY_MGR = "commissarymgr@marinationmobile.com";
var LOCATION_MGR = "amazonmanager@marinationmobile.com";
var TEMPLATE_NAME = "TEMPLATE** DUPLICATE ME";
var INVENTORY_ROWS = 39;
var TOTAL_ROWS = 55;

/**---INIT---*************************************************************************************/

/* @OnlyCurrentDoc */
var SS = SpreadsheetApp.getActiveSpreadsheet();

/**---FUNCTIONS---*********************************************************************************/

function onOpen(e){
  var UI = SpreadsheetApp.getUi();
  UI.createMenu('Marination')
    .addItem('Duplicate Template', 'cloneTemplate') //function to add
  .addSeparator()
  .addItem('Order Active Sheet', 'orderSheet') //function to add
    .addToUi();
  
  try{
    var lastSheet = SS.getSheets()[ss.getSheets().length - 1];
    if(lastSheet.isSheetHidden()) lastSheet.showSheet();
    SS.setActiveSheet(lastSheet);
  }
  catch(e){
    var functionName = "onOpen";
    errorReport(getDate() + " Error=" + functionName + ", " + e.message); 
  }
  
  hideSheets();
}


function hideSheets(){
  try{
    var allSheets = SS.getSheets();
    allSheets[2].showSheet(); //make sure current sheets isn't hidden
    for(index = 0; index < allSheets.length; index++){ //hide all but last sheet
      if(!allSheets[index].isSheetHidden() && index != 2){
        allSheets[index].hideSheet();
      }
    }
  }
  catch(e){
    var functionName = "hideSheets";
    errorReport(getDate() + " Error=" + functionName + ", " + e.message); 
  }
}


function hideColumns(){
  try{
    var sheet = SS.getActiveSheet();
    if(sheet){
      var daySelect = null;
      var d = new Date();
      if(d.getHours() > 17) daySelect = getDay() + 1; //If time is past 5pm select next day
      else daySelect = getDay();
      hideColumnsSwitch(daySelect, sheet);
    }
    else errorReport(getDate() + " Error=hideColumns, could not find active sheet");
  }
  catch(e){
    var functionName = "hideColumns";
    errorReport(getDate() + " Error=" + functionName + ", " + e.message); 
  }
}

function showColumns(){
  try{
    var UI = SpreadsheetApp.getUi();
    var sheet = SS.getActiveSheet();
    sheet.showColumns(1,25);
    UI.alert("Warning: Sensitive information is displayed, edit carefully.");
  }
  catch(e){
    var functionName = "showColumns";
    errorReport(getDate() + " Error=" + functionName + ", " + e.message); 
  }
}


function cloneTemplate() {                                        //Runs every Thursday from 5-6pm
  try{
    var d = new Date();
    var month = d.getMonth() + 1;
    var day = d.getDate() + 1; //adding one day due to the load list being done at night now.
    var calculateDate = d.getTime() + 604800000;                    //adding 7 days
    var dEnd = new Date();
    dEnd.setTime(calculateDate);
    var monthEnd = dEnd.getMonth() + 1;
    var dayEnd = dEnd.getDate();
    var name = month + "." + day + " - " + monthEnd + "." + dayEnd; //Marination Current format
    var sheet = SS.getSheetByName(TEMPLATE_NAME).copyTo(SS);
    SpreadsheetApp.flush();
    sheet.setName(name);
    SS.setActiveSheet(sheet);
    SS.moveActiveSheet(3); //put sheet to top of list
    
    //copy the cell protection over from the template
    var rangeTop = SS.getRange('E1:Y4');
    rangeTop.protect().setDescription('Top').setWarningOnly(true);
    var rangeLeft = SS.getRange('A1:D'+ TOTAL_ROWS);
    rangeLeft.protect().setDescription('Left').setWarningOnly(true);
    var rangeTotals = SS.getRange('W5:Y'+ INVENTORY_ROWS);
    rangeTotals.protect().setDescription('Totals').setWarningOnly(true);
    var calc1 = INVENTORY_ROWS + 2;
    var calc2 = INVENTORY_ROWS + 3;
    var rangeCatering = SS.getRange('E' + calc1 + ':V' + calc2);
    rangeCatering.protect().setDescription('Catering Days').setWarningOnly(true);
    //gray columns protection
    var rangeSunday = SS.getRange('L5:L'+ INVENTORY_ROWS);
    rangeSunday.protect().setDescription('Sunday').setWarningOnly(true);
    var rangeMonday = SS.getRange('O5:O'+ INVENTORY_ROWS);
    rangeMonday.protect().setDescription('Monday').setWarningOnly(true);
    var rangeTuesday = SS.getRange('F5:F'+ INVENTORY_ROWS);
    rangeTuesday.protect().setDescription('Tuesday').setWarningOnly(true);
    var rangeWednesday = SS.getRange('U5:U'+ INVENTORY_ROWS);
    rangeWednesday.protect().setDescription('Wednesday').setWarningOnly(true);
    var rangeThursday = SS.getRange('F5:F'+ INVENTORY_ROWS);
    rangeThursday.protect().setDescription('Thursday').setWarningOnly(true);
    var rangeFriday = SS.getRange('I5:I'+ INVENTORY_ROWS);
    rangeFriday.protect().setDescription('Friday').setWarningOnly(true);
    
    if(sheet.isSheetHidden()){
      sheet.showSheet();
    }
  }
  catch(e){
    var functionName = "cloneTemplate";
    errorReport(getDate() + " Error=" + functionName + ", " + e.message); 
  }
  
  hideSheets();
  hideColumns();
}


function nightlyUpdate(){                                       //Runs every day from 5-6 pm
  try{
    var allSheets = SS.getSheets();
    var lastSheet = SS.getSheets()[SS.getSheets().length - 1];
    if(lastSheet.isSheetHidden()) lastSheet.showSheet();          //unhide current sheet
    lastSheet.showColumns(1,25);
    for(index = 0; index < allSheets.length - 1; index++){
      if(!allSheets[index].isSheetHidden()){
        allSheets[index].hideSheet();
      }
      //allSheets[index].showColumns(1,25);                         //run expand on all old sheets
    }
    hideColumnsSwitch(getDay() + 1 , lastSheet);
  }
  catch(e){
    var functionName = "nightlyUpdate";
    errorReport(getDate() + " Error=" + functionName + ", " + e.message); 
  }
}


function sendEmail(){
  try{
    var UI = SpreadsheetApp.getUi();
    var subject = "Amazon Load List Completed";
    var htmlBody = "<div class='emailContainer' style='display:block;margin:10px auto;padding:20px 0px;width:80%;background-color:#f6c805;'><table style='display:block;width:400px;padding:0px;margin:0px auto;border:0px;border-collapse:separate;'><tr><td id='marinationBanner' style='max-width:400px;text-align:center;padding:0px;border-radius:5px;border-color:#242a2b;background-color:white;border:0px;'><h1 style='margin:6px 0px 15px 0px;'>(｡◕‿◕｡)🍳Mahalo🌮(ง°ل͜°)ง</h1></td></tr><tr><td style='max-width:400px;text-align:center;padding:0px;border-radius:5px;border:solid4px;border-color:#242a2b;'><img style='display:block;width:100%;border-radius:3px;margin:0px;' id='marinationImage' src='https://instagram.fsnc1-1.fna.fbcdn.net/t51.2885-15/e35/16789740_1415581778516738_4003761117397516288_n.jpg'></td></tr></table></div>";
    //MailApp.sendEmail("klynicol@gmail.com", subject, null, {htmlBody:  htmlBody}); //testing
    MailApp.sendEmail(COMMISSARY_MGR, subject, null, {htmlBody:  htmlBody});
    MailApp.sendEmail(LOCATION_MGR, subject, null, {htmlBody:  htmlBody});
    UI.alert("Notification sent to: " + COMMISSARY_MGR  + " & " +  LOCATION_MGR);
  }
  catch(e){
    var functionName = "sendEmail";
    errorReport(getDate() + " Error=" + functionName + ", " + e.message); 
  }
}


function hideColumnsSwitch(day, sheet){
  try{
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
      default: // Default and Saturday
        sheet.hideColumns(1);
        sheet.hideColumns(23,3);
        break;
    }
  }
  catch(e){
    var functionName = "hideColumnsSwitch";
    errorReport(getDate() + " Error=" + functionName + ", " + e.message); 
  }
}

function orderSheet(){
  var UI = SpreadsheetApp.getUi();
  var position = UI.prompt("Set Position: ");
  SS.moveActiveSheet(Number(position.getResponseText()));
}


function getDay(){
 var d = new Date();
 return d.getDay(); //0 = Sunday
}


function getDate(){ //Date Stamp for Errors
 var d = new Date();
 return String(d.getMonth() + ", " + d.getDate());
}


function errorReport(error){
 MailApp.sendEmail("klynicol@gmail.com", "Amazon Script Error", error);
}

