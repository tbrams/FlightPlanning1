function onOpen() {
    SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .createMenu('Navigation')
      .addItem('Update METAR', 'getMETAR')
      .addItem('Start planning', 'showSidebar')
      .addToUi();
  
  check_timestamp();
 
  // Set STATUS fields to N/A
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Planning');
  sheet.getRange("STATUSLDG").setBackgroundRGB(228, 228, 228).setValue("N/A");
  sheet.getRange("STATUSTO").setBackgroundRGB(228, 228, 228).setValue("N/A");
  sheet.getRange("STATUSMB").setBackgroundRGB(228, 228, 228).setValue("N/A");

}

function showSidebar() {
var html = HtmlService
            .createTemplateFromFile('Sidebar')
            .evaluate();
  html.setTitle('Flight Planning')
           .setWidth(300);

  SpreadsheetApp.getUi().showSidebar(html);
}


/**
 * Automatically fetch the current METAR and TAFs from dmi.dk and extract selected specially
 * formatted values to be used for flight navigation planning.
 *
 * @param {none} 
 * @return nothing but builds a table including the following fields
 *   Airodrome ICAO ID
 *   Wind direction / Wind Strength
 *   QNH
 *   Temperature
 *   Timestamp
 *
 * @customfunction
 */
function getMETAR() {

  var queryString = Math.random();
  var cellFunction = '=IMPORTHTML("https://www.dmi.dk/vejr/i-luften/metar-og-taf/?' + queryString + '","table",1)';

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var METARSheet = spreadsheet.getSheetByName('METAR');
  //METARSheet.activate();

  // paste the result into this cell hardcoded (this is a page not intended for the user)
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('METAR').getRange('A1').setValue(cellFunction);

  // Make an "easier to digest" table with the details we need for planning
  // Again this is done in a hardcoded cell location
  list_airports();
}


/**
 * Check timestamp against current METAR and warn if expired
 *
 * @param  {none} 
 * @return {none} 
 *
 * @customfunction
 */
function check_timestamp() {  
  // Get the current time in Zulu format and fetch day of month, hours and minutes for validity check
  var formattedDate = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'");
  var zRegEx=new RegExp(/\-(\d\d)T(\d\d)\:(\d\d)/);
  var dom=parseInt(zRegEx.exec(formattedDate)[1],10);
  var hrs=parseInt(zRegEx.exec(formattedDate)[2],10);
  var min=parseInt(zRegEx.exec(formattedDate)[3],10);
  
  // Get the DEP and the ARR Airfields - if they are blank use the first available METAR, currently in M2 on the METAR sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('METAR');
  var depAF= ss.getRange("DEP").getValue();
  var arrAF= ss.getRange("ARR").getValue();
  Logger.log("depAF: "+depAF+" arrAF: "+arrAF);

  // get METAR timestamps for each of them
  
  var metarTime=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('METAR').getRange('M2').getValue();
  var Mdom=parseInt(metarTime.substring(0,2),10);
  var Mhrs=parseInt(metarTime.substring(2,4),10);
  var Mmin=parseInt(metarTime.substring(4,6),10);

  var ageInMinutes = (hrs-Mhrs)*60+min-Mmin;
  Logger.log("METAR age in minutes: "+ageInMinutes);
  var msg="METAR is no longer valid...";
  if (Mdom!=dom || ageInMinutes>30*4) {
    Browser.msgBox(msg, "Use the Navigation Menu to update the METAR data", Browser.Buttons.OK);
  }   
}



/**
 * Parse the METAR information and create a compact list with the relevant info starting under
 * Cell I1 on the METAR sheet. 
 *
 * The function requires RAW metar data to be present in cell A1 and down. 
 *
 * @customfunction
 *
 */
function list_airports() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("METAR");
  var data = sheet.getDataRange().getValues();
  var metararr = [{}];
  var row=0;
  for (var i = 0; i < data.length; i++) {
    
    var airport;
    var published;
    var wind;
    var temp;
    var qnh;
    
    var qRegex = new RegExp(/q(\d{4})/);
    var tRegex = new RegExp(/(m?\d\d)\/m?\d\d/);

    //  Check to make sure we are looking at the METAR data
    if (data[i][0]=="METAR" || data[i][0]=="SPECI") {
      row=row+1;
      // Then parse the raw data into airport, wind and QNH information 
      metararr=data[i][2].split(" ");
      // Airport will always be the first token here. The next one is the time of issue ... unless some of the text is "not available"
      airport = metararr[0].toUpperCase();
      Logger.log("Processing: "+airport);
      if (data[i][2].indexOf("not available")<0) {
        published = metararr[1];

        // Third (ordinal number 2) will be either "auto" or the wind info we are looking for
        if (metararr[2]=="auto")
           wind=metararr[3];
        else
          wind=metararr[2];
      
        // Format wind information. At this point it is either something like "29018kt", "25023g34kt" or even "vrb02kt"
        // First remove "kt" if that is present
        if (wind.indexOf("kt")>0) wind=wind.substring(0,wind.indexOf("kt"));

        // Isolate Wind direction and Wind strength
        var wd=wind.substring(0,3);
        wind=wind.substring(3,wind.length);
         
        // Now check for wind gust - if present add half of gust difference to steady wind
        if (wind.indexOf("g")>0) {        
          var w1=wind.substring(0, wind.indexOf("g"));
          var w2=wind.substring(wind.indexOf("g")+1, wind.length);
        
          wind = (parseInt(w1)+(parseInt(w2)-parseInt(w1))/2).toFixed(0);
        }
      
        // Finally assemble the wind info in a format we are familiar with
        wind=wd+"/"+wind;
 
        // Get the QNH and the temperature directly from the RAW data using Regular Expressions
        qnh=qRegex.exec(data[i][2])[1];
        Logger.log("data[i][2]): %s",data[i][2]);
       temp=tRegex.exec(data[i][2])[1];
       Logger.log("temp done");
              
        var cell =sheet.getRange("I1");
        cell = cell.offset(row,0); cell.setValue(airport);
        cell = cell.offset(0,1); cell.setValue(wind);
        cell = cell.offset(0,1); cell.setValue(qnh);
        cell = cell.offset(0,1); cell.setValue(temp);
        cell = cell.offset(0,1); cell.setValue(published);
      } else {
        var cell =sheet.getRange("I1");
        cell = cell.offset(row,0); cell.setValue(airport);
        cell = cell.offset(0,1); cell.setValue("na");
        cell = cell.offset(0,1); cell.setValue("na");
        cell = cell.offset(0,1); cell.setValue("na");        
        cell = cell.offset(0,1); cell.setValue("na");        
     }
    }
  }
}



/**
 * Do a VLOOKUP on the range provided. Check the index column against key value and
 * use offset to identify the column data to be returned
 *
 * @param{'ekrk'}  value    The value to look up.
 * @param {C2:L30} table    The data range to check.
 * @param {0}      index    Check this column in the table.
 * @param {2}      offset   The column with the result to return.
 *
 * @return         result   Matching value or nothing
 *
 * @customfunction
 */
function tbVLookUp(value, table, index, offset) {

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('METAR');
  var lastRow=sheet.getRange(table).getLastRow();
  var datatable=sheet.getRange(table).getValues();
  for (row=0;row<lastRow-1;row++) {
    if (datatable[row][index]==value) {
       return String(datatable[row][index+offset]);
    }
  }
}

/**
 * Find and return first available row after the data in a datarange
 *
 * @param {"Planning"}  folder  The sheet to use.
 * @param {"RouteTable"} table  datarange to check.
 * @param {0}            column Column to check.
 *
 * @return         result   Absolute address or -1 if full
 *
 * @customfunction
 */
function firstEmptyRow(folder, table, column) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(folder);
  first = sheet.getRange(table).getRowIndex()+1;
  last = sheet.getRange(table).getLastRow();
  for (ro=1;ro<last-first;ro++){
    if (sheet.getRange(table).offset(ro, 0).getValue()==""){
      return (ro+first-1);
    }
  }
  return -1;
}

function testFirstEmptyRow() {
  Logger.log(firstEmptyRow("Planning", "RouteTable", 0));
}

/**
 * Write the value of the text parameter in the cell address of the Navigation folder
 *
 * @param {"C2"}   cell     The cell address.
 * @param {"EKBI"} value    The value to write.
 *
 * @return         result   Matching value or nothing
 *
 * @customfunction
 */
function writeCell(cell, value) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Planning');
  sheet.getRange(cell).setValue(value);
}



/**
 * Write values in all the cells intended for Altitude info.
 *
 * We will assume the following rule for the average wind direction/strength
 * between departure and arrival Aerodrome: From 0 to 3000 ft, the wind will
 * veer towards +30 degrees and gain 50% more strength at 3000 ft. We assume 
 * this is a linear transition.
 *
 * @param {1200}   feet     The desired flight altitude
 *
 * @return         
 *
 * @customfunction
 */
function writeAltCells(feet) {
  folder="Planning";
  
  // Wind calculation
  // get wind direction and strength at both DEP and ARR
  [[WD1, WS1], [WD2, WS2]]=getWind();
  Logger.log("Wind info: "+WD1+", "+WS1+", "+WD2+", "+WS2);
  avgWindDir=(WD1+WD2)/2.;
  avgWindStr=(WS1+WS2)/2.;
  if (feet<3000) {
    // Calculate slightly veering and increasing wind towards 3000
    WDir=Math.round(avgWindDir+30/3000.*feet);
    WStr=Math.round(avgWindStr*(1+0.5/3000.*feet));
  } else {
    // Assube the wind has veered +30 degrees and increased 50% in strength
    WDir=avgWindDir+30;
    WStr=avgWindStr*1.5;
  }
  Logger.log("Calculated average alt wind: "+WDir+"/"+WStr);
  
  
  table="RouteTable";
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(folder);
  first = sheet.getRange(table).getRowIndex()+1;
  last=firstEmptyRow(folder, table, 0);
  Logger.log('First row: '+first);
  Logger.log('First empty: '+last);
  
  for (ro=1;ro<=last-first;ro++){
    if (sheet.getRange(table).getCell(ro+1, 1).getValue()!=""){
      // Populate the ALT cell in the RoutePlan for this row
      sheet.getRange(table).getCell(ro+1, 3).setValue(feet);
      // Populate the W/S cell in the RoutePlan for this row
      sheet.getRange(table).getCell(ro+1, 6).setValue(WDir+"/"+("00" + WStr).slice(-2));
    }
  }
}

function testWriteAltCells() {
   writeAltCells(2000);
}

/**
 * Create a new entry in the flight plan
 *
 * @param {"Kalred"}   name     Name of alternative destination
 * @param {"10"}       nm       Distance in Nautical Miles
 * @param {"283"}      tt       True Track to destination
 *
 * @return         
 *
 * @customfunction
 */
function updateRouteTable(name, nm, tt) {
  folder="Planning";
  table="RouteTable";
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(folder);
  last=firstEmptyRow(folder, table, 0);
  
  
  // First move the Alternate destination table with header and table header one down
  oldRow=sheet.getRange("ADErange").getRow()-3;
  dataRow=sheet.getRange("ADErange").getRow();
  oldBlock="A"+oldRow+":P"+dataRow;
  newDest="A"+(oldRow+1);
  
  
  sheet.getRange(oldBlock).moveTo(sheet.getRange(newDest));
  // Update the named range to new position
  newRange="A"+(dataRow+1)+":P"+(dataRow+1);
  var nrlist = sheet.getNamedRanges();
  for (var i = 0; i < nrlist.length; i++) {
     var nr = nrlist[i];
     if (nr.getName() == 'ADErange') {
        var range = sheet.getRange(newRange); // new Range
        nr.setRange(range);
     }
  }  
  
  
  // copy second row (because it has the accumulated time) to this empty space
  second = sheet.getRange(table).getRowIndex()+2;
  dest="A"+last;
  sourceRange="A"+second+":P"+second;
  sheet.getRange(sourceRange).copyTo(sheet.getRange(dest));
  
  // Insert data in the appropriate cells
  Distdest="A"+last;
  TTdest="H"+last;
  Namedest="M"+last;
  sheet.getRange(Distdest).setValue(nm);
  sheet.getRange(TTdest).setValue(tt);
  sheet.getRange(Namedest).setValue(name);
  
  
  // Finally copy the trip total time to the TimeTrip field in the FuelCalc table
  sheet.getRange("O"+last).copyTo(sheet.getRange("TimeTrip"), {contentsOnly:true});
}


function tesUpdateRouteTable() {
    updateRouteTable("GRA",117, 42);
}

/**
 * Populate the alternative destination cells
 *
 * @param {"Kalred"}   name     Name of alternative destination
 * @param {"10"}       nm       Distance in Nautical Miles
 * @param {"283"}      tt       True Track to destination
 *
 * @return         
 *
 * @customfunction
 */
function writeAltDestCells(name, nm, tt) {
  folder="Planning";
  table="RouteTable";
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(folder);
  first = sheet.getRange(table).getRowIndex()+1;
  last=firstEmptyRow(folder, table, 0);
 

 // Clear everything below the route table and 5 rows down
  range="A"+last+":P"+(last+4)
  sheet.getRange(range).clear();
  
 // Get and copy the header style from route table 
  dest="A"+(last+1);
  sheet.getRange(table).offset(-2,0).getCell(1,1).copyTo(sheet.getRange(dest));
  sheet.getRange(dest).setValue("Alternative Destination");


  // Copy title- and first-row from routeplanning to new alternative
  dest="A"+(last+3);
  headerRange="A"+(first-1)+":P"+(first);
  sheet.getRange(headerRange).copyTo(sheet.getRange(dest));

  // Insert data in the appropriate cells
  destRow=last+4;  
  Distdest="A"+destRow;
  TTdest="H"+destRow;
  Namedest="M"+destRow;
  sheet.getRange(Distdest).setValue(nm);
  sheet.getRange(TTdest).setValue(tt);
  sheet.getRange(Namedest).setValue(name);
  
  
  // Finally update the named range "ADErange" to refer to the data in this table
  newRange="A"+destRow+":P"+destRow;
  var nrlist = sheet.getNamedRanges();
  for (var i = 0; i < nrlist.length; i++) {
     var nr = nrlist[i];
     if (nr.getName() == 'ADErange') {
        var range = sheet.getRange(newRange); // new Range
        nr.setRange(range);
     }
  }  
}



function LookupCheat(value, table, offset) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  spreadsheet.getSheetByName('METAR').activate();
  
  var formula='VLOOKUP("'+value+'",'+table+','+offset+',0)';
  SpreadsheetApp.getActiveSheet().getRange('HiddenCell').setFormula(formula);

  Logger.log("Result in hidden cell is: "+SpreadsheetApp.getActiveSheet().getRange('HiddenCell').getValue());

}

/**
 * For the purpose of best practise and avoid having to type all this again and again
 * this function will include the requested file.
 *
 * As an alternative, we could write something like
 * <?!= HtmlService.createHtmlOutputFromFile('stylesheet').getContent(); ?>
 *
 * But this is shorter and more elegant
 *
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

/**
 * Convert feet to meter
 *
 * @param  {Value|Range} feet The distance in feet 
 * @return {Value}            The distance in meter
 *
 * @customFunction
*/
function f2m(feet){
   return feet/3.28;
}

/**
 * Convert kg to lbs
 *
 * @param  {Value|Range} kg   The weight in kg 
 * @return {Value}            The weight in lbs
 *
 * @customFunction
*/
function k2p(kg){
   return kg*2.2;
}


/**
 * Calculate and return wind components as a range
 * 
 * @param {1} AD  Aerodrome selector: 
 *                1: takeoff, 
 *                2: landing
 *
 * @return  result   [[WC, SWC]]
 *
 * @customfunction
 */
function calculateWC(AD) {
  folder="Planning";
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(folder);
  // Get a range with RW strings
  RWs = sheet.getRange("RWrange").getValues();
  
  // Make sure there are no letters on the RW, for example like in "22R"
  RW1=RWs[0][0];
  RW2=RWs[0][1];  
  if ((typeof RW1)=="string") RW1=RW1.replace(/[a-zA-Z]$/, '');
  if ((typeof RW2)=="string") RW2=RW2.replace(/[a-zA-Z]$/, '');
 
  Logger.log("RW1: "+RW1);
  Logger.log("RW2: "+RW2);
  
  // Get the wind for both DEP and ARR
  [[W1Dir, W1Str], [W2Dir, W2Str]]=getWind();
  
  // Calculate the angle between RW and wind
  Ang1=W1Dir-RW1*10;
  Ang2=W2Dir-RW2*10;  
  
  Logger.log("Ang1: "+Ang1);
  Logger.log("Ang2: "+Ang2);
  
  // Calculate and return WC and SWC for either DEP or ARR in a range
  var res = [];
  if (AD==1)
     res.push([W1Str*Math.round(100*Math.cos(Ang1*Math.PI/180))/100, W1Str*Math.round(100*Math.sin(Ang1*Math.PI/180))/100]);
  else
     res.push([W2Str*Math.round(100*Math.cos(Ang2*Math.PI/180))/100, W2Str*Math.round(100*Math.sin(Ang2*Math.PI/180))/100]);

  Logger.log(res);
  
  return res;
}

/**
 * Get wind information for both DEP and ARR and return a range with parsed info
 *
 * @return  result   [[WDir1, WStr1],[WDir2, WStr2]]
 *
 * @customfunction
 */
function getWind() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Planning");
  WVs = sheet.getRange("WVrange").getValues();
  
  WV1=WVs[0][0];
  [W1Dir,W1Str]=WV1.split("/");
  Logger.log("W1Dir: "+W1Dir);
  Logger.log("W1Str: "+W1Str);
  
  WV2=WVs[0][1];
  [W2Dir, W2Str]=WV2.split("/");
  Logger.log("W2Dir: "+W2Dir);
  Logger.log("W2Str: "+W2Str);

  res=[];
  res.push([parseInt(W1Dir,10), parseInt(W1Str,10)]);
  res.push([parseInt(W2Dir,10), parseInt(W2Str,10)]);
  return res;
}

function testGetWind() {
  Logger.log("GetWind: "+getWind());
}

function testCalculateWC() {
  Logger.log("calculateWC: "+calculateWC(1));
}



/**
 * Get information for both DEP and ARR and return a range with parsed info
 *
 * @return  result   [DEP,ARR]
 *
 * @customfunction
 */
function getADlist() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Planning");
  DEP = sheet.getRange("DEP").getValue();
  ARR = sheet.getRange("ARR").getValue();
  return [DEP, ARR];
  
}
