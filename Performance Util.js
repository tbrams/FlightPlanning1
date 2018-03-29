/**
 * Find the interpolated result by looking up in the table passed
 *
 * @param {Val|Range}  Temp       The temperature in C
 * @param {Val|Range}   PA        The Pressure Altitude of the airfield
 * @param {"DEP"|"ARR"} table     If DEP will select the TO Performance Table, otherwise the LDG... 
 *
 * @return {Date} The new date.
 * @return {Value} Best matching result based on table or error code.
 *       possible values:
 *         `-1` temperature outside table range
 *         `-2` pressure altitude outside table range
 *         `-3` Both of these outside
 *
 * @customfunction
 */
function interpolate(t, pa, table) {
  const DEP=0;
  const ARR=1;
  
  if (table=="DEP") 
    tableRange="TO_Distances";
  else if (table=="ARR")
    tableRange="LD_Distances";
  else  
    tableRange=table;
  
  var table = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Performance').getRange(tableRange).getValues();
  ncols = table[0].length;
  nrows = table.length;
    
   // First check if the parameters can be used in the table
   ErrorCode=0;
   if (t<table[1][0] || t>table[nrows-1][0]) {
      Logger.log("Temperature unusable");
      ErrorCode=-1;
   }
   if (pa<table[0][1] || t>table[0][ncols-1]) {
      Logger.log("pa unusable");
      ErrorCode=ErrorCode-2;
   }
   
   // Return -1 on temp error, -2 on pa error, -3 on both
   if (ErrorCode<0) return;

   // Check PA against values in first row
   var indexPA=table[0].indexOf(pa);
   
   // find matching columns
   indexPA1=indexPA2=0;
   if (indexPA!=-1) {
      indexPA1=indexPA2=indexPA;
   } else {
      for (col=ncols; col>0; col--) {
         if (table[0][col-1]<pa) {
            indexPA1=col-1;
            indexPA2=indexPA1+1;
            break;
         }
      }
      Logger.log("Will use PA from "+table[0][indexPA1]+" to "+table[0][indexPA2]);
   }

   // Check temperature against first column in table
   var indexT=(table.map(function(value,index){ return value[0];})).indexOf(t);

   // Find matching rows
   indexT1=indexT2=0;
   if (indexT!=-1) {
      indexT1=indexT2=indexT;
   } else {
     for (row=nrows; row>0; row--) {
         if (table[row-1][0]<t) {
            indexT1=row-1;
            indexT2=indexT1+1;
            break;
         }
      }
      Logger.log("Will use temp from "+table[indexT1][0]+" to "+table[indexT2][0]);
   }

   PA11=table[indexT1][indexPA1];
   PA12=table[indexT1][indexPA2];
   PA21=table[indexT2][indexPA1];
   PA22=table[indexT2][indexPA2];
   Logger.log('PA11: '+PA11);
   Logger.log('PA12: '+PA12);
   Logger.log('PA21: '+PA21);
   Logger.log('PA22: '+PA22);
   
   PA0=table[0][indexPA1];
   T0=table[indexT1][0];
   
   // for t1
   I1=PA11;
   I2=PA21;   
   if (indexPA1!=indexPA2) {
      Logger.log('diff: '+(table[0][indexPA2]-table[0][indexPA1]));
      Logger.log('indexT1 '+indexT1);
      Logger.log('indexPA1 '+indexPA1);
      Logger.log('table[]: '+table[indexT1][indexPA1]);
      Logger.log('pa: '+pa+' '+(PA11));
      I1=PA11+(PA12-PA11)/(table[0][indexPA2]-table[0][indexPA1])*(pa-PA0);
      I2=PA21+(PA22-PA21)/(table[0][indexPA2]-table[0][indexPA1])*(pa-PA0);
   }
   Logger.log("I1: "+I1);
   Logger.log("I2: "+I2);
   
   I=I1;
   if (indexT1!=indexT2) {
      I=I+(I2-I1)/(table[indexT2][0]-table[indexT1][0])*(t-T0);
   }
   Logger.log("I: "+I);
   
   return I;
}


function testInterpolate(){
   interpolate(35,3000, 'TO_Distances');
}


/**
 * Find the engine specific metrics by looking up on RPM and return as array
 *
 * @param {2450}  rpm       Rotations per minute
 *
 * @return {Range} Best matching result based on table or error code
 * 
 *         `-1` ERROR, RPM outside table range - otherwise
 *         [`BHP`, `TAS`, `FF`] 
 *
 * @customfunction
 */
function getEngineStats(rpm) {
  tableRange="EnginePerformance";
   var table = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Performance').getRange(tableRange).getValues();
   ncols = table[0].length;
   nrows = table.length;
 
  
   // First check if the parameters can be used in the table
   if (rpm>table[0][0] || rpm<table[nrows-1][0]) {
      Logger.log("RPM not covered by table");
      return -1; 
   }
   
   // Check if the RPM value can be looked up directly 
   var indexRPM=(table.map(function(value,index){ return value[0];})).indexOf(rpm);
   Logger.log("indexRPM is "+indexRPM);

   indexR1=indexR2=0;
   if (indexRPM!=-1) {
      // if the value is in the table set both pointers at this value
      indexR1=indexR2=indexRPM;
   } else {
   // Find matching rows
     for (row=0; row<nrows; row++) {
         if (table[row][0]<rpm) {
            indexR1=row;
            indexR2=indexR1-1;
            break;
         }
      }
      Logger.log("Will use rpm from "+table[indexR1][0]+" to "+table[indexR2][0]);     
   }

  // Values for interpolation. Note that I will not use the "one down" method
  // here, because it is a lot easier to just subtract 3% in final TAS

  BHP1 = table[indexR1][1];
  BHP2 = table[indexR2][1];
  TAS1 = table[indexR1][3];
  TAS2 = table[indexR2][3];
  FF1 = table[indexR1][5];
  FF2 = table[indexR2][5];
     
   R1=table[indexR1][0];
   R2=table[indexR2][0];
   BHP=BHP1;
   TAS=TAS1;
   FF=FF1;
   if (indexR1!=indexR2) {
      BHP=BHP+(BHP2-BHP1)/(R2-R1)*(rpm-R1);
      TAS=TAS+(TAS2-TAS1)/(R2-R1)*(rpm-R1);
      FF=FF+(FF2-FF1)/(R2-R1)*(rpm-R1);
   }
   Logger.log("BHP: "+BHP);
   Logger.log("TAS: "+TAS);
   TAS=TAS*.974;
   Logger.log("Reduced TAS: "+TAS);
   Logger.log("FF: "+FF);

  return [BHP, TAS, FF];
}


function testEngineStats(){
  Logger.log("Enginestats: "+getEngineStats(2301));
}


/**
 * Write values in all the cells intended for RPM info
 * That is BHP, TAS and FF ... 
 *
 * @param {2550}   rpm     The RPM value.
 *
 * @return         
 *
 * @customfunction
 */
function writeEngineCells(rpm) {
  folder="Planning";
  table="FuelCalcTable";
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(folder);

  results=getEngineStats(rpm);
  
  // In this array we have: BHP, RPM and then FF
  sheet.getRange(table).getCell(2, 2).setValue(results[0]); // BHP
  sheet.getRange(table).getCell(4, 2).setValue(results[0]);
  sheet.getRange(table).getCell(5, 2).setValue(results[0]);  

  sheet.getRange(table).getCell(2, 3).setValue(results[2]); // FF
  sheet.getRange(table).getCell(4, 3).setValue(results[2]);
  sheet.getRange(table).getCell(5, 3).setValue(results[2]);  

  sheet.getRange(table).getCell(2, 6).setValue(rpm); 
  sheet.getRange(table).getCell(4, 6).setValue(rpm);
  sheet.getRange(table).getCell(5, 6).setValue(rpm);  

  // Now fill in the TAS values in the Flight Plan table
  table="RouteTable";
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(folder);
  first = sheet.getRange(table).getRowIndex()+1;
  last=firstEmptyRow(folder, table, 0);
  Logger.log('First row: '+first);
  Logger.log('First empty: '+last);
  
  for (ro=1;ro<=last-first;ro++){
    if (sheet.getRange(table).getCell(ro+1, 1).getValue()!=""){
      sheet.getRange(table).getCell(ro+1, 5).setValue(results[1]);
    }
  }
  
}


/**
 * Write the value in all the cells intended for RPM info
 *
 * @param {2550}   rpm     The RPM value.
 *
 * @return         
 *
 * @customfunction
 */
function writeRPMCells(rpm) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Planning');
  sheet.getRange("FuelCalcTable").getCell(2, 8).setValue(rpm);
  sheet.getRange("FuelCalcTable").getCell(4, 8).setValue(rpm);
  sheet.getRange("FuelCalcTable").getCell(5, 8).setValue(rpm);  
}
