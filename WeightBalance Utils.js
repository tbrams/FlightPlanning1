/**
 * Populate the Mass fields in the spreadsheet based on these parameters
 *
 * @param{'C2'}    cell     The cell address.
 * @param {"EKBI"} value    The value to write.
 *
 * @return         result   Matching value or nothing
 *
 * @customfunction
 */
function writeMassFields(W_Front, W_Back, W_Lug_Front, W_Fuel_Total, W_Fuel_Spent) {
    writeCell('W_Front', W_Front);
    writeCell('W_Back', W_Back);
    writeCell('W_Lug_Front', W_Lug_Front);
    writeCell('W_Fuel_Total', W_Fuel_Total);
    writeCell('W_Fuel_Spent', W_Fuel_Spent);
}

function testWriteMasses(){
  writeMassFields(100,101,102,103,104);
}


/**
 * Get the start and landing weight and moment figures from spreadsheet and return as an
 * array with two ranges (also arrays).
 *
 */
function getNumbers(){
   return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('WeightBalance').getRange('STW').getValues()
          .concat(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('WeightBalance').getRange('LDW').getValues());
}


/**
 * Get Flight registration and type from Weighttable into an array and return a html select statement
 * with all this
 *
 */
function flightSelectionHTML() {
   tableRange="WeightTable";
   table = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Performance').getRange(tableRange).getValues();

  ncols=table[0].length;
  var HTML="";
  for (i=1;i<ncols;i++) {
    if (table[i][0]=="") break;
    HTML+='<option value="'+table[i][1]+'">'+table[i][0]+'</option>';
  }
  Logger.log(HTML);
  return HTML;
  
}

function testFS(){
   Logger.log(flightSelectionHTML());
}



/**
 * Procedure used internally to convert coordinates in string format from the envelope table into
 * an array of 2D coordinates.
 *
 * @return         result   Array of 2D coordinates,  or -1 for no envelope found
 *
 * @customfunction
 */
function getEnvelopeCoords() {
   AircraftType = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("WeightBalance").getRange("OYType").getValue();
   EnvelopeTable = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("WeightBalance").getRange("EnvelopeTable").getValues();

   // Go throught the entries in the envelope table until the type is found, then get the string with coordinates
   var cordString="";  
   nrows=EnvelopeTable.length;
   for (i=1;i<nrows;i++) {
     if (EnvelopeTable[i][0]==AircraftType){
        cordString=EnvelopeTable[i][1];
        break;
      }
    
      // Just in case - no matching type, return error code
      if (EnvelopeTable[i][0]=="") {
        Logger.log("Error: No envelope available for "+AircraftType);
        return -1;
      }
   }  
  
   // Convert the string with coordinates into an array of x,y coordinates
   cordString=cordString.trim();  // clean leading and trailing spaces 
   parts=cordString.split(",");
   // Clean the broken apart pieces and make it an array of [x, y] coordinates
   coords=[];
   for (i=0; i<parts.length;i=i+2){
      // clean again, in case there are leading or trailing spaces
      parts[i]=parts[i].trim();
      // Then remove brackets
      x=parts[i].substr(1,parts[i].length-1);
      y=parts[i+1].substr(0, parts[i+1].length-1);   
      coords.push([x, y]);
   }
  
   return coords;
}


function testGetEnvelopeCoords() {
  result=getEnvelopeCoords();
  Logger.log("Envelope coords");
  for (i in result) {
    Logger.log(i+"): "+result[i]);
  }
}


/**
 * Get and convert envelope coords to 4D by adding extra "N" values to all existing pairs.
 * Subsequently append the two singular points for Take Off and Landing and return the 
 * whole thing as an array.
 *
 * @param {Int}    TOW    Take off weight.
 * @param {Int}    LDW    Landing Weight.
 * @param {Int}    TOM    Take off Moment.
 * @param {Int}    LDM    Landing Moment.
 *
 * @return         result   Array of 4D coordinates,  or -1 for no envelope found
 *
 * @customfunction
 */
function patchEnvelopeData(TOW, LDW, TOM, LDM) {
   var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Planning');
  
   coords=getEnvelopeCoords();  
   if (coords!=-1) {
    
     // Before pathing up with additional mnulls for the graphics, check if the points
     // are inside the envelope and update the STATUSMB field on the front page accordingly
     if ((inside(TOM, TOW, coords)) && (inside(LDM, LDW, coords))) {
        // Mark it green
        sheet.getRange("STATUSMB").setBackgroundRGB(0, 255, 0).setValue("Yes");
     } else {
        // mark it red
        sheet.getRange("STATUSMB").setBackgroundRGB(255, 0, 0).setValue("No");    
     } 
    
     // Patch all coordinates with extra nulls to make them 4D
     newCoords=[];
     for (i in coords) {
         newCoords.push([coords[i][0],coords[i][1], "N", "N"]);   
     }
  
     // Finally append the two singular points
     newCoords.push([TOM, "N", TOW, "N"]);
     newCoords.push([LDM, "N", "N", LDW]);

     return newCoords;
    
   } else {
     // mark status field it yellow
     sheet.getRange("STATUSMB").setBackgroundRGB(255, 255, 0).setValue("No");    
     // And pass on the error code 
     return -1;
   }
}

function testPatchEnvelopeData(){
  result=patchEnvelopeData(2200, 2100, 100, 90);
  Logger.log("Patched Envelope");
  for (i in result) {
    Logger.log(i+"): "+result[i]);
  }
  
}


/**
 *
 * Check if a point is inside a polygon
 *
 * @param {5}            x        The X coordinate of the point
 * @param {3}            y        The Y coordinate of the point
 * @param {[[x,y]...]}   shape    An array of points representing the
 *                                corners in the shape.
 *
 * @return boolean                True if point is inside shape
 *
 * @customfunction
 */

function inside(x,y, shape) {
   flag=false;
   var yi;
   x1=parseInt(shape[0][0]);
   y1=parseInt(shape[0][1]);
   for (i=1; i<shape.length; i++) {
      x2=parseInt(shape[i][0]);
      y2=parseInt(shape[i][1]);
      if (x2!=x1){
        yi=(y2-y1)/(x2-x1)*(x-x1)+y1;
        if (y<=yi) flag=!flag;
      } else if (y2==y1) {
        flag=!flag;
      };
      [x1,y1]=[x2,y2];    
   }

   return flag;
}


function testInside(){
   var MyPoly=[[50, 1420],[82, 2140],[98, 2340],[112, 2340],[62, 1300],[50, 1300],[50, 1420]];
   if (inside(88,2166, MyPoly))
      Logger.log("Point is inside");
    else
      Logger.log("Point is outside");
}
