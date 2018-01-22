/*************************************************************************************************

    ---------Marination Mobile LLC, Load List Script---------
    --------------Created by Mark Wickline 2018--------------
    ----------------------Ma Kai Version---------------------
    
**---EDIT VARIABLES---****************************************************************************/

var COMMISSARY_MGR = "commissarymgr@marinationmobile.com"; //edit inside quotations
var LOCATION_MGR = "";                                     //edit inside quotations
var TEMPLATE_NAME = "TEMPLATE ** DUPLICATE ME";            //edit inside quotations
var INVENTORY_ROWS = 35;                                   //no quotations needed
var TOTAL_ROWS = 35;                                       //no quotations needed

var PRINT_RANGE = "B2:Z35";                                //edit inside quotations
var PRINTER_EMAIL = "alkimakai@hpeprint.com";              //edit inside quotations

var ADMIN_EMAIL = "klynicol@gmail.com";                    //edit inside quotations, error reporting and testing email

/**---INIT---*************************************************************************************/

/* @OnlyCurrentDoc */                             //Helps limit permissions requested when using this script
var SS = SpreadsheetApp.getActiveSpreadsheet();   //GLOBAL VARIABLE... An Object of the SpreadsheetApp class, Holds current Spreadsheet

/**---FUNCTIONS---*********************************************************************************/

function onOpen(e){
  var UI = SpreadsheetApp.getUi();                   //Get the User Interface from SpreadSheet App
  UI.createMenu('Marination')                        //Adding a new menu to the top called Marination
    .addItem('Duplicate Template', 'cloneTemplate')  //add cloneTemplate function from this script to the Marination Menu
  .addSeparator()                                    //-------------------
  .addItem('Print Template', 'printTemplate')        //add printTemplate function from this script to the Marination Menu
  .addSeparator()                                    //-------------------
  .addItem('Order Active Sheet', 'orderSheet')       //add orderSheet function from this script to the Marination Menu
    .addToUi();
  
  try{
      hideSheets();                   //runs the hideSheets function from this script
  }
  catch(e){                           //See errorReport function at the end of this script.
    var functionName = "onOpen";
    errorReport(getDate() + " Error=" + functionName + ", " + e.message); 
  }
}


function hideSheets(){
  try{
    var allSheets = SS.getSheets();                            //getting an array that holds all the sheets, Hidden and Unhidden
    var currentSheet = allSheets[2];                           //setting the current sheet to a variable called currentSheet
    if(currentSheet.isSheetHidden()) currentSheet.showSheet(); //make sure current sheet isn't hidden
    SS.setActiveSheet(currentSheet);                           //set the current sheet to be the active sheet
    
    /* --- Hide all Sheets Accept the 3rd sheet in the Array (The Current Sheet) ---- */
    for(index = 0; index < allSheets.length; index++){ 
      if(!allSheets[index].isSheetHidden() && index != 2){
        allSheets[index].hideSheet();
      }
    }
  }
  catch(e){                           //See errorReport function at the end of this script.
    var functionName = "hideSheets";
    errorReport(getDate() + " Error=" + functionName + ", " + e.message); 
  }
}


function hideColumns(){
  try{
    var sheet = SS.getActiveSheet();                   //Gets the current sheet being viewed
    if(sheet){                                         //If a sheet is being viewed execute this block. Mostly for error purposes
      var daySelect = getDay();                        //Calling function getDay from this script and setting it to a variable, will return a number 0-6
      var d = new Date();                              //Setting a variable for the current time to check current user hour
      if(d.getHours() >= 17) daySelect += 1;           //If time is past 5pm select next day
      hideColumnsSwitch(daySelect, sheet);             //Runs the hideColumnsSwitch function from this script
    }
    else errorReport(getDate() + " Error=hideColumns, could not find active sheet");
  }
  catch(e){                           //See errorReport function at the end of this script.
    var functionName = "hideColumns";
    errorReport(getDate() + " Error=" + functionName + ", " + e.message); 
  }
}

function showColumns(){
  try{
    var UI = SpreadsheetApp.getUi();                                            //Get the User Interface from SpreadSheet App, for use with UI.Alert below
    var sheet = SS.getActiveSheet();                                            //Get the active sheet from the spreadsheet object
    sheet.showColumns(1,30);                                                    //Show all columns on that sheet <- Edit the second number for the number of rows on the sheet
    UI.alert("Warning: Sensitive information is displayed, edit carefully.");   //Show a warning to the user
  }
  catch(e){
    var functionName = "showColumns";
    errorReport(getDate() + " Error=" + functionName + ", " + e.message); 
  }
}


function cloneTemplate() {                                          //Runs every Thursday from 5-6pm
  try{
    /*Building the name for the new sheet (Example: "1.21 - 1.27")*/
    var d = new Date();
    var month = d.getMonth() + 1;                                   //adding 1 because of Javascript format, changing to standard months 1-12 instead of 0-11
    var day = d.getDate() + 1;                                      //adding one day due to the load list being done at night now.
    var calculateDate = d.getTime() + 604800000;                    //adding 7 days, one extra then we need because were adding one day to the day variable.
    var dEnd = new Date();
    dEnd.setTime(calculateDate);
    var monthEnd = dEnd.getMonth() + 1;
    var dayEnd = dEnd.getDate();
    var name = month + "." + day + " - " + monthEnd + "." + dayEnd; //Setting the finished name to a variable called name
    
    /*Dulpicating the template, setting it's name to the one we just built, and setting it to the active sheet*/
    var sheet = SS.getSheetByName(TEMPLATE_NAME).copyTo(SS);        //Getting the template and coppying it to the SpreadSheet App object (the real duplicate happens here)
    SpreadsheetApp.flush();                                         //Flushing the App of any pending scripts needing to be run, mostly preventative action
    sheet.setName(name);                                            //Setting the new sheet name to the current Marination format
    SS.setActiveSheet(sheet);
    SS.moveActiveSheet(3);                                          //Put sheet to top of list, just behind the template
    
    /*----------------------------------------------------------------------
    -----------------PROTECTING RANGES WITH A WARNING-----------------------
    THIS FUNCTIONALITY CANNOT HAPPEN WHILE MANUALLY DUPLICATING THE TEMPLATE
    ----------------------------------------------------------------------*/
    var rangeTop = SS.getRange('E1:AD3');
    rangeTop.protect().setDescription('Top').setWarningOnly(true);
    
    var rangeLeft = SS.getRange('A1:F'+ TOTAL_ROWS);
    rangeLeft.protect().setDescription('Left').setWarningOnly(true);
    
    var rangeTotals = SS.getRange('AB3:AD'+ INVENTORY_ROWS);
    rangeTotals.protect().setDescription('Totals').setWarningOnly(true);
    
    //var calc1 = INVENTORY_ROWS + 2;
    //var calc2 = INVENTORY_ROWS + 3;
    //var rangeCatering = SS.getRange('E' + calc1 + ':V' + calc2);
    //rangeCatering.protect().setDescription('Catering Days').setWarningOnly(true);
    
    //gray columns protection
    var rangeSunday = SS.getRange('Q3:Q'+ INVENTORY_ROWS);
    rangeSunday.protect().setDescription('Sunday').setWarningOnly(true);
    
    var rangeMonday = SS.getRange('T3:T'+ INVENTORY_ROWS);
    rangeMonday.protect().setDescription('Monday').setWarningOnly(true);
    
    var rangeTuesday = SS.getRange('W3:W'+ INVENTORY_ROWS);
    rangeTuesday.protect().setDescription('Tuesday').setWarningOnly(true);
    
    var rangeWednesday = SS.getRange('Z3:Z'+ INVENTORY_ROWS);
    rangeWednesday.protect().setDescription('Wednesday').setWarningOnly(true);
    
    var rangeThursday = SS.getRange('H3:H'+ INVENTORY_ROWS);
    rangeThursday.protect().setDescription('Thursday').setWarningOnly(true);
    
    var rangeFriday = SS.getRange('K3:K'+ INVENTORY_ROWS);
    rangeFriday.protect().setDescription('Friday').setWarningOnly(true);
    
    var rangeSaturday = SS.getRange('N3:N'+ INVENTORY_ROWS);
    rangeSaturday.protect().setDescription('Saturday').setWarningOnly(true);
    
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


function nightlyUpdate(){ //Runs every day from 5-6 pm, sets the active sheet to the correct sheet and switches the displayed column to the next day.
  try{
    var currentSheet = SS.getSheets()[2];
    currentSheet.showSheet();
    SS.setActiveSheet(currentSheet);
    currentSheet.showColumns(1,30); //unhide everything
    hideColumns();
  }
  catch(e){
    var functionName = "nightlyUpdate";
    errorReport(getDate() + " Error=" + functionName + ", " + e.message); 
  }
}


/* VALUE KEY FOR COLUMNS

A=1 G=7  M=13 S=19 Y=25  AE=31
B=2 H=8  N=14 T=20 Z=26  AF=32
C=3 I=9  O=15 U=21 AA=27
D=4 J=10 P=16 V=22 AB=28
E=5 K=11 Q=17 W=23 AC=29
F=6 L=12 R=18 X=24 AD=30 

*/

function hideColumnsSwitch(day, sheet){
  try{
    switch (day){
      case 0: //Sunday
        sheet.hideColumns(1);     //If only 1 number is in the parentheses, only that column gets hidden.
        sheet.hideColumns(7,9);   //If two numbers are in parentheses, first number is the column to start hiding, second number is the amount of columns to hide after that column.
        sheet.hideColumns(19,12); //^^ Use the key above for converting columns to numbers.
        break;
      case 1: //Monday
        sheet.hideColumns(1);
        sheet.hideColumns(7,12);
        sheet.hideColumns(22,9);
        break;
      case 2: //Tuesday
        sheet.hideColumns(1);
        sheet.hideColumns(7,15);
        sheet.hideColumns(25,6);
        break;
      case 3: //Wednesday
        sheet.hideColumns(1);
        sheet.hideColumns(5,18);
        sheet.hideColumns(28,3);
        break;
      case 4: //Thursday
        sheet.hideColumns(1);
        sheet.hideColumns(10,21);
        break;
      case 5: //Friday
        sheet.hideColumns(1);
        sheet.hideColumns(7,3);
        sheet.hideColumns(13,18);
        break;
      case 6: // Saturday
        sheet.hideColumns(1);
        sheet.hideColumns(7,6);
        sheet.hideColumns(16,15);
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

function printTemplate() {
  try{
    var sheet = SS.getSheetByName(TEMPLATE_NAME); //getting the template
    if(sheet.isSheetHidden()) sheet.showSheet();
    SS.setActiveSheet(sheet);
    
/* VALUE KEY FOR COLUMNS

A=1 G=7  M=13 S=19 Y=25  AE=31
B=2 H=8  N=14 T=20 Z=26  AF=32
C=3 I=9  O=15 U=21 AA=27
D=4 J=10 P=16 V=22 AB=28
E=5 K=11 Q=17 W=23 AC=29
F=6 L=12 R=18 X=24 AD=30 

*/
    
    /* hide and show columns */
    sheet.hideColumns(1);
    //sheet.hideColumns(6); //Summer pars
    sheet.hideColumns(9);
    sheet.hideColumns(12);
    sheet.hideColumns(15);
    sheet.hideColumns(18);
    sheet.hideColumns(21);
    sheet.hideColumns(24);    //If only 1 number is in the parentheses, only that column gets hidden.
    sheet.hideColumns(27, 4); //If two numbers are in parentheses, first number is the column to start hiding, second number is the amount of columns to hide after that column.
    sheet.showColumns(7, 2);  //^^ Use the key above for converting columns to numbers.
    sheet.showColumns(10, 2);
    sheet.showColumns(13, 2);
    sheet.showColumns(16, 2);
    sheet.showColumns(19, 2);
    sheet.showColumns(22, 2);
    sheet.showColumns(25, 2);
    
    /* Helping with Authorization */
    var forScope = DriveApp.getRootFolder();
    
    /* Get Authorization code and sheet ID */
    var token = ScriptApp.getOAuthToken();  //Auth token
    var gid = sheet.getSheetId();
    
    
    var url = 'https://docs.google.com/spreadsheets/d/'+ss.getId()+'/export?exportFormat=pdf&format=pdf' // export as pdf / csv / xls / xlsx
    + '&size=0'                            // paper size legal / letter / A4
    + '&range=' + PRINT_RANGE              // cell range
    + '&portrait=false'                    // orientation, false for landscape
    + '&fitw=true'                         // fit to page width, false for actual size
    + '&sheetnames=false&printtitle=false' // hide optional headers and footers
    + '&pagenumbers=false&gridlines=true'  // hide page numbers and gridlines
    + '&fzr=false'                         // do not repeat row headers (frozen rows) on each page
    + '&gid='+gid;                         // the sheet's Id
    
    var attachment = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' +  token
      }
    });
    
    //MailApp.sendEmail(PINTER_EMAIL, 'Ma Kai Load List', null , {attachments: [attachment.getBlob().getAs(MimeType.PDF)]});
    MailApp.sendEmail(ADMIN_EMAIL, 'Ma Kai Load List', null , {attachments: [attachment.getBlob().getAs(MimeType.PDF)]}); //testing
  }
  catch(e){
    var functionName = "printTemplate";
    Logger.log(getDate() + " Error=" + functionName + ", " + e.message); 
  } 

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
 MailApp.sendEmail(ADMIN_EMAIL, "Ma Kai Script Error", error); //send email with the error code
}
