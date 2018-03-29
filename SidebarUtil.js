/**
 * Create list content from a data range in a spreadsheet
 * @param {string} spreadsheet The name of the spreadsheet with the source data.
 * @param {Range} dataRange The range of name of the datarange to use as source.
 *
 * @return {String} a HTML formatted string with values for the select box
 *
 * @customFunction
 */
function readIntoList(spreadsheet, dataRange) {
  const ICAO=0;
  const FULLNAME=4;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(spreadsheet);
  var listInfo = sheet.getRange(dataRange).getValues();
  
  // Make sure we do not get duplicates into the drop down list by checking for ICAO code
  var unique = listInfo.reduce(
        function(a,b){
          if ((a.map(function(value,index) {return value[ICAO]}).indexOf(b[ICAO]) < 0) ) a.push(b);
          return a;
        },[]);
  
  // Keep the ICAO code for Value and the Full name for the user option.
  var HTML='';
  for (i in unique) {
    HTML+='<option value="'+unique[i][ICAO]+'">'+unique[i][FULLNAME]+'</option>';
  }
  Logger.log('TOB: '+HTML);
  
  return HTML;
}

/**
 * Create list with detail content from a data range in a spreadsheet. If provided with two
 * column arguments, the second will be used as the VAL in the OPTION statements. If no secondary
 * column is provided, the VAL will be a sequential number starting at 1
 *
 * @param {string} spreadsheet The name of the spreadsheet with the source data.
 * @param {Range}  dataRange   The range of name of the datarange to use as db.
 * @param {String} key         The lookup key for the db.
 * @param {Int}    Column      The data column from DB to use as result.
 * @param {Int}    Column2     Optional: The secondary data column from DB to use as result.
 *
 * @return {String} a HTML formatted string with values for the select box
 *
 * @customFunction
 */
function getDetails(spreadsheet, dataRange, key, column, column2) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(spreadsheet);
  var myDB = sheet.getRange(dataRange).getValues();
  var myList=[];
  for (i in myDB) 
     if (myDB[i][0]==key) myList.push(myDB[i]);

  // Create a HTML formatted list of values
  var HTML='';

  for (i in myList) {
    if (column2) ind=myList[i][column2]; else ind=i+1;
    HTML+='<option value="'+ind+'">'+myList[i][column]+'</option>';
  }
  Logger.log('TOB: '+HTML);
  
  return HTML;
}

/**
 *
 * Calculate take off correction from percentage to meter and write this in the appropriate row
 *
 * @param {15} percentage  The percentage to use.
 * @param {4}  factor      The dominant factor (1..7)
 *
 * @customFunction
 */
function writeTOfactor(percentage, factor) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Start/Ldg');
  var HandBook=sheet.getRange('STAHB').getValue();
  
  // Clear all previous factors by making a range covering all factors first
  myStart=sheet.getRange("STAK1").getA1Notation();  // this is in the F12 notation
  myEnd=parseInt(myStart.slice(1))+6;               // last row 
  myRange=myStart+":"+myStart.substr(0,1)+myEnd;
  sheet.getRange(myRange).clear();


  if (factor!=0) {  
     myRange="STAK"+factor;
     sheet.getRange(myRange).setValue(HandBook*percentage/100);
  
     Logger.log('Compensation factor: '+factor+' is '+HandBook*percentage/100+'m');
  }
  
  // Then calculate wind factors 
  wcArr=calculateWC(1);
  Logger.log("wcArr: "+wcArr);
  wc=wcArr[0][0];
  Logger.log("wc: "+wc);
  
  if (wc<0) {  // tail wind 
     sheet.getRange("STAW1").offset(0,-2).setValue(wc);
     sheet.getRange("STAW1").setValue(1.5*Math.abs(wc)*0.04*HandBook);
     sheet.getRange("STAW2").clear();
     sheet.getRange("STAW2").offset(0,-2).clear();
  } else {
     sheet.getRange("STAW2").offset(0,-2).setValue(wc);
     sheet.getRange("STAW2").setValue(-0.01*wc/2*HandBook);
     sheet.getRange("STAW1").clear();
     sheet.getRange("STAW1").offset(0,-2).clear();
  }
  
  // Finally set status flag for Take off distance on front page
  minTO=sheet.getRange("STAMIN").getValue();
  AD=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Planning').getRange("DEP").getValue();
  RW=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Planning').getRange("DEP_RW").getValue();
  avlTO=getRWL(AD,RW);
  sheet.getRange("STAMIN").offset(1,0).setValue(avlTO);  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Planning');
  if (avlTO>=minTO) {
    sheet.getRange("STATUSTO").setBackgroundRGB(0, 255, 0).setValue("Yes");
  } else {
    sheet.getRange("STATUSTO").setBackgroundRGB(255, 0, 0).setValue("No");    
  }
}

function testWriteTOfactor() {
   writeTOfactor(5, 1);
}



/**
 *
 * Calculate landing correction from percentage to meter and write this in the appropriate row
 *
 * @param {15} percentage  The percentage to use.
 * @param {4}  factor      The dominant factor (1..4)
 *
 * @customFunction
 */
function writeLDfactor(percentage, factor) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Start/Ldg');
  var HandBook=sheet.getRange('LANHB').getValue();
  
  // Clear all previous factors by making a range covering all factors first
  myStart=sheet.getRange("LANK1").getA1Notation();  // this is in the F12 notation
  myEnd=parseInt(myStart.slice(1))+4;               // last row 
  myRange=myStart+":"+myStart.substr(0,1)+myEnd;
  sheet.getRange(myRange).clear();
  
  if (factor!=0) {
     myRange="LANK"+factor;
     sheet.getRange(myRange).setValue(HandBook*percentage/100);
  
     Logger.log('Compensation factor: '+factor+' is '+HandBook*percentage/100+'m');
  }
  
  
  // Finally set status flag for Take off distance on front page
  minLDG=sheet.getRange("LANMIN").getValue();
  AD=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Planning').getRange("ARR").getValue();
  RW=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Planning').getRange("ARR_RW").getValue();
  avlLDG=getRWL(AD,RW);
  sheet.getRange("LANMIN").offset(1,0).setValue(avlLDG);  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Planning');
  if (avlLDG>=minLDG) {
    sheet.getRange("STATUSLDG").setBackgroundRGB(0, 255, 0).setValue("Yes");
  } else {
    sheet.getRange("STATUSLDG").setBackgroundRGB(255, 0, 0).setValue("No");    
  }
}

function testWriteLDfactor() {
   writeLDfactor(5, 4);
}


/**
 *
 * Return runway length
 *
 * @param {"EKRK"} AD  ICAO AD Name.
 * @param {11}     RW  Runway
 *
 * @customFunction
 */

function getRWL(AD,RW) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Aerodromes");
  var myDB = sheet.getRange("AerodromeData").getValues();
  rwl=0;
  for (i in myDB) {
    if (myDB[i][0]==AD && myDB[i][1]==RW) {
         rwl=myDB[i][2];
         break;
    }
  }
  
  Logger.log('Length '+rwl);
  return rwl;
}


/**
 *
 * Just for testing the ability to write to a cell from here
 *
 * @customFunction
 */
function writeResult(text) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('METAR');
  var mycell=sheet.getRange('Z10');
  mycell.setValue(text);
  Logger.log('Result put in cell: '+text);
}

