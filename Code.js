function createCustomMenu() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Check Authorization before using!')
      .addItem('Check Authorization and Scripts', 'authorize')
    .addToUi();
}

//MAIN SHEET AND OTHERS
SpreadsheetApp.getActiveSpreadsheet().getSheet
const sheetNameCBS1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Bus Passenger Details CEBU1")
const tripCountSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TripCounter");
const sorterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sorter")
const sheetNameMonthly = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Monthly Summary")
const sheetNamePWD = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PWD")
const sheetNameSenior = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Senior Citizens")
const sheetNameOverall = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Overall")
const infoSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Information Sheet")
const ncrMainMonitoringSpreadsheet = SpreadsheetApp.openById("1oleMJSy-HsZn23ulUI51OiWAMZdqbOmMumy3qOYVDqs").getSheetByName("RECORDS : CBS1");

//FOLDERS
const dmrFile = DriveApp.getFileById('1ixnavfOqvjPxtGpCW2X6wmQTmnbCvyuPJ0jNLIVO2xU');
const monthlyDestFolder = DriveApp.getFolderById("1KivTYs2E7ReQiSjgxgbI--gsYIQdtv06"); 
const dailyDestFolder = DriveApp.getFolderById("1yuR8PliOKMzDp3_Rr8QsQ5ZJRjBsHZuz"); 
const dmrDestFolder = DriveApp.getFolderById("15xniPYp-oGUiElmIAOt7eKDMfbT2l9ku"); 

//Sheet Data .JSON
var configFile = DriveApp.getFileById('1HSfMZmtUrQ8GVdFvQCh7-cjHpUNF3qQD');
var configContent = configFile.getBlob().getDataAsString();
// Parse the JSON content into an object
var configData = JSON.parse(configContent);
// Access the configuration data
const sheetRangesOverall = configData.sheetRangesOverall;
const sheetRangesSenior = configData.sheetRangesSenior;
const sheetRangesPWD = configData.sheetRangesPWD;

//LASTROW
const lastrow_CBS1 = sheetNameCBS1.getLastRow();
const lastrow_tripCounter = tripCountSheet.getLastRow();
const lastrow_monthly = sheetNameMonthly.getLastRow();
const lastrow_pwd = sheetNamePWD.getLastRow();
const lastrow_senior = sheetNameSenior.getLastRow();
const lastrow_overall = sheetNameOverall.getLastRow();
const lastrow_mainMonitoring = ncrMainMonitoringSpreadsheet.getLastRow();

//DATE AND OTHERS
const values = sheetNameCBS1.getRange(lastrow_CBS1, 1, 1, 5).getValues()[0];
const date = Utilities.formatDate(new Date(), "UTC+8", "MM/dd/yyyy");
const addresetCounter = tripCountSheet.getRange("A2").getValue();
const tripNumberTotal = tripCountSheet.getRange("B2").getValue();

//DATA
const dateData = sheetNameCBS1.getRange(lastrow_CBS1, 1);
const tripNumberData = sheetNameCBS1.getRange(lastrow_CBS1, 2);
const ageData = sheetNameCBS1.getRange(lastrow_CBS1, 3);
const genderData = sheetNameCBS1.getRange(lastrow_CBS1, 4);
const pwdData = sheetNameCBS1.getRange(lastrow_CBS1, 5)

// Ranges
const totalPassengerRange = sheetNameOverall.getRange("J8");
const totalMaleMonthlyRange = sheetNameMonthly.getRange("X10");
const totalFemaleMonthlyRange = sheetNameMonthly.getRange("X11");
const totalPassengerMonthlyRange = sheetNameMonthly.getRange("U9:AA9");
const grandTotalTripsMonthly = sheetNameMonthly.getRange("G9:M9");
const totalSeniorRange = sheetNameSenior.getRange("J8");
const totalSeniorMonthlyRange = sheetNameMonthly.getRange("X13");
const totalPWDRange = sheetNamePWD.getRange("J8");
const totalPWDMonthlyRange = sheetNameMonthly.getRange("X12");
const counterAM = sorterSheet.getRange("A15");
const counterPM = sorterSheet.getRange("A17");
const tripNumberToday = infoSheet.getRange("D6");
const infosheetMale = infoSheet.getRange("D7")
const infosheetFemale = infoSheet.getRange("D8")
const infosheetPWD = infoSheet.getRange("D9");
const infosheetSenior = infoSheet.getRange("D10");
const totalPassengersAM = infoSheet.getRange("D11");
const totalPassengersPM = infoSheet.getRange("D12");
const malePassengersAM = infoSheet.getRange("D13");
const malePassengersPM = infoSheet.getRange("D14");
const femalePassengersAM = infoSheet.getRange("D15");
const femalePassengersPM = infoSheet.getRange("D16");
const tripTimeAM = infoSheet.getRange("D17");
const tripTimePM = infoSheet.getRange("D18");
const infosheetTotal = infoSheet.getRange("D19");
const grandTotalTripsRange = infoSheet.getRange("D26")
const busDriver = infoSheet.getRange("D28");
const busConductor = infoSheet.getRange("D29");
const plateNumber = infoSheet.getRange("D30");
const documentCounterDaily = sorterSheet.getRange("A13");
const documentNumberDaily = documentCounterDaily.getValue();
const documentCounterMonthly = sorterSheet.getRange("A19");
const dailyResetCopyLink = sorterSheet.getRange("A21")
const dailyMonitoringReportLink = sorterSheet.getRange("A23")
const ovpTemplateMonthlySheet = sheetNameMonthly.getRange("B2:AA5")
const ovpTemplateOverallSheet = sheetNameOverall.getRange("B2:AH4")
const ovpTemplateSeniorSheet = sheetNameSenior.getRange("B2:AH4")
const ovpTemplatePWDSheet = sheetNamePWD.getRange("B2:AH4")
const documentNumberMonthly = documentCounterMonthly.getValue();
const totalPassenger = 0;
const totalPassengerMonthly = 0;
const amPMTripsRange = tripTimeAM.getValue() + tripTimePM
const amPMTrips = tripTimeAM.getValue() + tripTimePM.getValue();

// Define a queue to store data rows
var dataQueue = [];

//ROW OFFSET FOR SORTER SHEET
var rowNumber = sorterSheet.getRange("A9").getValue();
const rowOffset_sheets = rowNumber
const count_sheets = sheetNameOverall.getRange("C12:AH42").getValues().flat().filter(String).length;
const lastrow_sheets = count_sheets + rowOffset_sheets;

var rowNumberMonthly = sorterSheet.getRange("A11").getValue();
const rowOffset_monthly = rowNumberMonthly;
const count_monthly = sheetNameMonthly.getRange("B47:AA47").getValues().flat().filter(String).length;
const lastrowro_monthly = count_monthly + rowOffset_monthly;

function test(){
}

// Function to batch update values in ranges
function batchUpdateValues(rangeValuePairs) {
  rangeValuePairs.forEach(([range, value]) => {
    range.setValue(value);
  });
}

function enqueueData(rowData) {
    dataQueue.push(rowData);
    Logger.log('Data enqueued for processing: ' + JSON.stringify(rowData));
    // Trigger the queue processing
    processQueue();
}

function processQueue() {
    // Check if the queue is not empty
    while (dataQueue.length > 0) {
        // Dequeue the first item from the queue
        var rowData = dataQueue.shift();
        Logger.log('Data dequeued for processing: ' + JSON.stringify(rowData));
        
        // Process the dequeued data row
        processPassengerData(rowData);
    }
}

function incrementCounts(totalRange, monthlyRange, monthlyMaleRange, monthlyFemaleRange, gender) {
  const isMale = gender === 'Male';
  const isFemale = gender === 'Female';
  
  if (isMale) {
    monthlyMaleRange.setValue(monthlyMaleRange.getValue() + 1);
  } else if (isFemale) {
    monthlyFemaleRange.setValue(monthlyFemaleRange.getValue() + 1);
  }

  const totalPassengerValue = totalPassengerRange.getValue() + 1;
    totalRange.setValue(totalPassengerValue);
    totalPassengerRange.setValue(totalPassengerValue);

  const totalPassengerMonthlyValue = totalPassengerMonthlyRange.getValue() + 1;
    monthlyRange.setValue(totalPassengerMonthlyValue);
    totalPassengerMonthlyRange.setValue(totalPassengerMonthlyValue);
}

function sortGenderLogic(currentTripNumber, currentGender, formattedTime) {
  if (!(currentTripNumber in sheetRangesOverall)) {
    return;
  }

  const genderKey = currentGender === 'Male' ? 'male' : 'female';
  const genderRange = sorterSheet.getRange(sheetRangesOverall[currentTripNumber][genderKey]);
  const addGender = genderRange.getValue() + 1;
  genderRange.setValue(parseInt(addGender));

  let passengerRange;

  if (formattedTime === 'AM') {
    passengerRange = currentGender === 'Male' ? malePassengersAM : femalePassengersAM;
  } else {
    passengerRange = currentGender === 'Male' ? malePassengersPM : femalePassengersPM;
  }

  const totalPassengersRange = formattedTime === 'AM' ? totalPassengersAM : totalPassengersPM;

  // Create an array to hold all the updates
  const updates = [
    [passengerRange, passengerRange.getValue() + 1], // Update passenger range based on gender
    [totalPassengersRange, totalPassengersRange.getValue() + 1], // Update total passengers range
  ];

  // Batch update all ranges
  batchUpdateValues(updates);
}

function sortAgeLogic(currentTripNumber, currentAge, currentGender) {
  if (!(currentTripNumber in sheetRangesSenior) || currentAge < 60) {
    return;
  }

  const genderKey = currentGender === 'Male' ? 'male' : 'female';
  const genderRange = sorterSheet.getRange(sheetRangesSenior[currentTripNumber][genderKey]);
  const addGender = genderRange.getValue() + 1;
  genderRange.setValue(parseInt(addGender));

  // Create an array to hold all the updates
  const updates = [
    [totalSeniorRange, totalSeniorRange.getValue() + 1], // Update totalSeniorRange
    [totalSeniorMonthlyRange, totalSeniorMonthlyRange.getValue() + 1], // Update totalSeniorMonthlyRange
    [infosheetSenior, infosheetSenior.getValue() + 1], // Update infosheetTotal
  ];

  // Batch update all ranges
  batchUpdateValues(updates);
}

function sortPWDLogic(currentTripNumber, currentPWD, currentGender) {
  if (!(currentTripNumber in sheetRangesPWD) || currentPWD !== 'Yes') {
    return;
  }

  const genderKey = currentGender === 'Male' ? 'male' : 'female';
  const genderRange = sorterSheet.getRange(sheetRangesPWD[currentTripNumber][genderKey]);
  const addGender = genderRange.getValue() + 1;
  genderRange.setValue(parseInt(addGender));

  // Create an array to hold all the updates
  const updates = [
    [totalPWDRange, totalPWDRange.getValue() + 1], // Update totalPWDRange
    [totalPWDMonthlyRange, totalPWDMonthlyRange.getValue() + 1], // Update totalPWDMonthlyRange
    [infosheetPWD, infosheetPWD.getValue() + 1], // Update infosheetTotal
  ];

  // Batch update all ranges
  batchUpdateValues(updates);
}

function doIt(isDeletedColumn) {
    Logger.log('doIt() function called.');

    var data = sheetNameCBS1.getRange("A" + lastrow_CBS1 + ":E" + sheetNameCBS1.getLastRow()).getValues();

    if (!isDeletedColumn && data.length > 0) {
        data.forEach(rowData => {
            enqueueData(rowData);
        });
        Logger.log('Data enqueued for processing.');
    } else if (isDeletedColumn) {
        var deletedTimestamp = isDeletedColumn.getTime();
        undoIncrements(deletedTimestamp);
        Logger.log('Deletion through AppSheet: Timestamp ' + deletedTimestamp + ' removed from processing.');
    } else {
        Logger.log('No data found in "Bus Passenger Details CBS1" sheet. Skipping data processing.');
    }
}

function processPassengerData(rowData) {
  var date = new Date(); // Replace this with your actual date
  Logger.log('Processing passenger data: ' + date);
  var timeZone = Session.getScriptTimeZone();
  var format = "a";
  var formattedTime = Utilities.formatDate(date, timeZone, format);
  Logger.log('Formatted Time: ' + formattedTime);
  
  var tripNumber = rowData[1];
  var age = rowData[2];
  var gender = rowData[3];
  var pwdStatus = rowData[4]; // Extract PWD status

  if (gender === 'Male' || gender === 'Female') {
    incrementCounts(totalPassengerRange, totalPassengerMonthlyRange, totalMaleMonthlyRange, totalFemaleMonthlyRange, gender);
  }

  sortGenderLogic(tripNumber, gender, formattedTime);
  sortAgeLogic(tripNumber, age, gender);
  sortPWDLogic(tripNumber, pwdStatus, gender); // Pass PWD status
}

function tripTimeStart(){
   var date = new Date(); // Replace this with your actual date
    Logger.log(date)
    var timeZone = Session.getScriptTimeZone();
    var format = "a";
  var formattedTime = Utilities.formatDate(date, timeZone, format);
  var valueTripStartAM = infoSheet.getRange("D17").getValue();
  var valueTripStartPM = infoSheet.getRange("D18").getValue();
  var tripGrandTotalRange = grandTotalTripsRange.getValue() + 1;
  var monthlyTotalTripsRange = grandTotalTripsMonthly.getValue() + 1;
  
    if (valueTripStartAM <= 0 && valueTripStartPM <= 0){
       if(formattedTime == "AM"){
        var tripTimeRangeAM = tripTimeAM.getValue() + 1;
        tripTimeAM.setValue(parseInt(tripTimeRangeAM));

      } else if (formattedTime == "PM"){
        var tripTimeRangePM = tripTimePM.getValue() + 1;
        tripTimePM.setValue(parseInt(tripTimeRangePM))
      }
    }
    grandTotalTripsRange.setValue(parseInt(tripGrandTotalRange));
    grandTotalTripsMonthly.setValue(parseInt(monthlyTotalTripsRange))
}

function mainFunction(e) {
    Logger.log('mainFunction() started.');
    Logger.log(JSON.stringify(e));

    // Define the column index you want to ignore (Column F is index 6)
    const ignoredColumnIndex = 6;

    if (e.changeType === 'EDIT' && e.source.getSheetName() === 'Bus Passenger Details CBS1') {
        var range = e.source.getActiveRange();

        // Check if the edited column index is the ignored column index
        if (range.getColumn() !== ignoredColumnIndex && range.getRow() > 1) {
            var deletedValue = range.getValue();
            var isDeletedColumn = deletedValue === '' ? sheetNameCBS1.getRange(range.getRow(), 1).getValue() : null;

            doIt(isDeletedColumn);
            tripTimeStart();
        }
    }

    Logger.log('mainFunction() completed.');
}

function triggerFunction() {
  // Check if it's the first day of the month
  var date = new Date();
  var isFirstDayOfMonth = date.getDate() == 1;
  Logger.log(isFirstDayOfMonth);

  if (isFirstDayOfMonth) {
    Logger.log("End Month Function is running.")
    endMonth();  
  } else {
    Logger.log("End Day Function is running.")
    endDay();
  }
}

function endDay() {
var date = new Date();
var isFirstDayOfMonth = date.getDate() == 1;
if (isFirstDayOfMonth) {
    endMonth();
  } else {

  resetDailySetTables(function() {
    setDateOverall(function() {
        setRouteforToday(function() {
          setTotalofTrips(function() {
              dailyMonitoringReport(function() {
                resetDailyGenerateReports(function() {
            });
          });
        });
      });
    });
  });
 }
}

function endMonth() {
  resetMonthSetTables(function() {
    setDateOverall(function() {
        setRouteforToday(function() {
          setTotalofTrips(function() {
             dailyMonitoringReport(function() {
              resetMonthGenerateReports(function() {
                  setRouteforTheDay(function() {
            });
          });
        });
      });
    });
  });
 });
}

// Date for Overall Sheet, PWD, and Senior / Grand Total of Trips.
function setDateOverall(callback){
  const rowOffset = 12
  const count = sheetNameOverall.getRange("B12:B42").getValues().flat().filter(String).length;
  const lastrow = count + rowOffset;
  sheetNameOverall.getRange(lastrow, 2).setValue(date)
  sheetNameSenior.getRange(lastrow, 2).setValue(date)
  sheetNamePWD.getRange(lastrow, 2).setValue(date)
  callback()
}

function tripTimeSort(e){
  if (e.source.getSheetName() === "TripCounter"){
  Logger.log(e.source.getSheetName())
  var values = tripCountSheet.getRange("B2").getValue();
  var date = new Date(); // Replace this with your actual date
  Logger.log(date)
  var timeZone = Session.getScriptTimeZone();
  var format = "a";
  var formattedTime = Utilities.formatDate(date, timeZone, format);
  Logger.log(formattedTime)
    if (values >= 2){
      if(formattedTime == "AM"){
        var tripTimeRangeAM = tripTimeAM.getValue() + 1;
        tripTimeAM.setValue(parseInt(tripTimeRangeAM))
        var tripNumberTotalToday = tripNumberToday.getValue() + 1;
        tripNumberToday.setValue(parseInt(tripNumberTotalToday))

      } else if (formattedTime == "PM"){
        var tripTimeRangePM = tripTimePM.getValue() + 1;
        tripTimePM.setValue(parseInt(tripTimeRangePM))
        var tripNumberTotalToday = tripNumberToday.getValue() + 1;
        tripNumberToday.setValue(parseInt(tripNumberTotalToday))
      }
     }
    } 
  }

function setRouteforTheDay(){
  // Get the current date
  var currentDate = new Date();
  
  // Array of weekday names
  var weekdays = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  
  // Get the day of the week as a number (0 for Sunday, 1 for Monday, and so on)
  var dayNumber = currentDate.getDay();
  
  // Get the day name from the array using the day number
  var dayName = weekdays[dayNumber];
  Logger.log(currentDate)
  Logger.log(dayNumber)
  Logger.log(dayName)

var timeZone = Session.getScriptTimeZone();
var format = "EEEE, dd MMMM yyyy";
var formatCheck = "EEEE"
var formattedDate = Utilities.formatDate(currentDate, timeZone, format);
Logger.log("Formatted Date: " + formattedDate);

    var destination = "";
  if (dayName == "Monday" || dayName == "Tuesday") {
    destination = "MANDAUE CITY ROUTE";
  } else if (dayName == "Wednesday" || dayName == "Thursday") {
    destination = "CEBU CITY ROUTE";
  } else if (dayName == "Friday" || dayName == "Saturday") {
    destination = "LAPU-LAPU CITY ROUTE";
  } else if (dayName == "Sunday") {
    destination = "NO TRIPS ON SUNDAYS";
  }

  tripTimeAM.setValue(parseInt(0))
  infoSheet.getRange("D5").setValue(destination);
  infoSheet.getRange("D2").setValue(formattedDate)

}

function setRouteforToday(callback) {
  const rowOffset_overall = 12
  const count_overall = sheetNameOverall.getRange("AH12:AH42").getValues().flat().filter(String).length;
  const lastrowro_overall = count_overall + rowOffset_overall;

  const rowOffset_senior = 12
  const count_senior = sheetNameSenior.getRange("AH12:AH42").getValues().flat().filter(String).length;
  const lastrowro_senior = count_senior + rowOffset_senior;

  const rowOffset_pwd = 12
  const count_pwd = sheetNamePWD.getRange("AH12:AH42").getValues().flat().filter(String).length;
  const lastrowro_pwd = count_pwd + rowOffset_pwd;


  // Get the current date
  var currentDate = new Date();
  
  // Array of weekday names
  var weekdays = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  
  // Get the day of the week as a number (0 for Sunday, 1 for Monday, and so on)
  var dayNumber = currentDate.getDay();
  
  // Get the day name from the array using the day number
  var dayName = weekdays[dayNumber];
  Logger.log(currentDate)
  Logger.log(dayNumber)
  Logger.log(dayName)

    var destination = "";
    var destinationMonthly = "";
  if (dayName == "Monday" || dayName == "Tuesday") {
    destination = "MANDAUE CITY ROUTE";
    destinationMonthly = "Mandaue City Route: Jagobiao - A.C. Cortes - B.B. Cabahug - A. Del Rosario - A.C Cortes - M.C. Briones";
  } else if (dayName == "Wednesday" || dayName == "Thursday") {
    destination = "CEBU CITY ROUTE";
    destinationMonthly = "Cebu City Route: SM Seaside - N. Bacalso - Osmena Blvd. - Capitol - IT Park - Escario St.- Capitol - Osmena Blvd., N. Bacalso - SM Seaside";
  } else if (dayName == "Friday" || dayName == "Saturday") {
    destination = "LAPU-LAPU CITY ROUTE";
    destinationMonthly = "Lapu-Lapu City Route: Mandaue City - Brgy. Mepz - Brgy Saac - Mactan, Lapu-Lapu City - Brgy. Soong - Brgy Maribago - Brgy. Marigondon - Cordova - Brgy Opon";
  } else if (dayName == "Sunday") {
    destination = "NO TRIPS ON SUNDAYS";
    destinationMonthly = "NO TRIPS ON SUNDAYS";
  }

  sheetNameMonthly.getRange(lastrowro_monthly, 2).setValue(date)
  sheetNameMonthly.getRange(lastrowro_monthly, 7).setValue(destinationMonthly);
  sheetNameOverall.getRange(lastrowro_overall, 34).setValue(destination);
  sheetNameSenior.getRange(lastrowro_senior, 34).setValue(destination);
  sheetNamePWD.getRange(lastrowro_pwd, 34).setValue(destination);
  var totalTripCount = sheetNameMonthly.getRange(lastrowro_monthly, 3).getValue() + amPMTrips;
  sheetNameMonthly.getRange(lastrowro_monthly, 3).setValue(parseInt(totalTripCount));
  sheetNameOverall.getRange(lastrowro_overall, 33).setValue(parseInt(amPMTrips));
  sheetNameSenior.getRange(lastrowro_senior, 33).setValue(parseInt(amPMTrips));
  sheetNamePWD.getRange(lastrowro_pwd, 33).setValue(parseInt(amPMTrips));
  callback();
}

function setTotalofTrips(callback){
var grandTotalTrips = grandTotalTripsMonthly.getValue() + amPMTrips;
grandTotalTripsMonthly.setValue(parseInt(grandTotalTrips));
callback()
}

function moveToCancelled() {
  Logger.log("moveToCancelled() started.");
  var date = new Date(); // Replace this with your actual date
  Logger.log(date)
  var timeZone = Session.getScriptTimeZone();
  var format = "a";
  var formattedTime = Utilities.formatDate(date, timeZone, format);
  Logger.log(formattedTime)

  var triggerCancel = "Deleted";
  var newTriggerCancel = "Deleted and Recorded";
  var red = "#FF0000";
  var values = sheetNameCBS1.getRange(2, 1, sheetNameCBS1.getLastRow(), 6).getValues();
  Logger.log("Values from CBS1 sheet: " + JSON.stringify(values));

  var { v, cells, rbRanges, rbRows } = values.reduce((o, r, i) => {
     var tripNumber = r[1];
    if (r[5] == triggerCancel) {
      Logger.log("Processing entry in row " + (i + 2) + " marked as Deleted");
      o.v.push([r[2], r[3], r[4]]);
      o.cells.push(`R${i + 2}`);
      o.rbRanges.push(`A${i + 2}:T${i + 2}`);
      o.rbRows.push(i + 2);

      Logger.log(`Processing entry in row ${i + 2} marked as Deleted`);
      Logger.log(`Entry Details: Trip Number: ${tripNumber}, Gender: ${r[3]}, Age: ${r[2]}, PWD: ${r[4]}`);

      if (r[3] === 'Male') {
        Logger.log("Processing Male entry");
        Logger.log(`Gender: ${r[3]}`);
        //Total Monthly X10
        var totalMaleMonthly = totalMaleMonthlyRange.getValue() - 1;
        totalMaleMonthlyRange.setValue(parseInt(totalMaleMonthly));
        //Total Monthly U9:AA9
        var totalPassengersMonthly = totalPassengerMonthlyRange.getValue() - 1;
        totalPassengerMonthlyRange.setValue(parseInt(totalPassengersMonthly))
        //Total Overall J8
        var totalPassenger = totalPassengerRange.getValue() - 1;
        totalPassengerRange.setValue(parseInt(totalPassenger));
        // Sorter Sheet
        var overallMaleRange = sorterSheet.getRange(sheetRangesOverall[tripNumber].male);
        var overallMaleDecrement = overallMaleRange.getValue() - 1;
        overallMaleRange.setValue(parseInt(overallMaleDecrement));

        //infosheet Passengers Male AM, Male PM, and Total Passengers AM
        if(formattedTime == "AM"){
        var malePassengerAM = malePassengersAM.getValue() - 1;
        malePassengersAM.setValue(parseInt(malePassengerAM))
        var totalPassengersAMValue = totalPassengersAM.getValue() - 1;
        totalPassengersAM.setValue(parseInt(totalPassengersAMValue))

      } else if (formattedTime == "PM"){
        var malePassengerPM = malePassengersPM.getValue() - 1;
        malePassengersPM.setValue(parseInt(malePassengerPM))
        var totalPassengersPMValue = totalPassengersPM.getValue() - 1;
        totalPassengersPM.setValue(parseInt(totalPassengersPMValue))
      }

      } else if (r[3] === 'Female') {
        Logger.log("Processing Female entry");
        Logger.log(`Gender: ${r[3]}`);
        //Total Monthly X11
        var totalFemaleMonthly = totalFemaleMonthlyRange.getValue() - 1;
        totalFemaleMonthlyRange.setValue(parseInt(totalFemaleMonthly));
        //Total Monthly U9:AA9
        var totalPassengerMonthly = totalPassengerMonthlyRange.getValue() - 1;
        totalPassengerMonthlyRange.setValue(parseInt(totalPassengerMonthly));
        //Total Overall J8
        var totalPassenger = totalPassengerRange.getValue() - 1;
        totalPassengerRange.setValue(parseInt(totalPassenger));
        //SorterSheet
        var overallFemaleRange = sorterSheet.getRange(sheetRangesOverall[tripNumber].female);
        var overallFemaleDecrement = overallFemaleRange.getValue() - 1;
        overallFemaleRange.setValue(parseInt(overallFemaleDecrement));

        //infosheet Passengers Female AM, Female PM, and Total Passengers AM
        if(formattedTime == "AM"){
        var femalePassengerAM = femalePassengersAM.getValue() - 1;
        femalePassengersAM.setValue(parseInt(femalePassengerAM))
        var totalPassengersAMValue = totalPassengersAM.getValue() - 1;
        totalPassengersAM.setValue(parseInt(totalPassengersAMValue))

      } else if (formattedTime == "PM"){
        var femalePassengerPM = femalePassengersPM.getValue() - 1;
        femalePassengersPM.setValue(parseInt(femalePassengerPM))
        var totalPassengersPMValue = totalPassengersPM.getValue() - 1;
        totalPassengersPM.setValue(parseInt(totalPassengersPMValue))
      }
      }

      if (r[2] >= 60) {
        if (r[3] === 'Male') {
          Logger.log("Processing Senior Male entry");
          Logger.log(`Gender: ${r[3]}`);
          Logger.log(`Age: ${r[2]}`);
          //Total Senior J8
        var totalSenior = totalSeniorRange.getValue() - 1;
        totalSeniorRange.setValue(parseInt(totalSenior));
        //Total Senior X13
        var totalSeniorMonthly = totalSeniorMonthlyRange.getValue() - 1;
        totalSeniorMonthlyRange.setValue(parseInt(totalSeniorMonthly));
        //SorterSheet
        var seniorMaleRange = sorterSheet.getRange(sheetRangesSenior[tripNumber].male);
        var seniorMaleDecrement = seniorMaleRange.getValue() - 1;
        seniorMaleRange.setValue(parseInt(seniorMaleDecrement));
        var infosheetSenior = infoSheet.getRange("D10")
        infosheetSenior.setValue(parseInt(infosheetSenior.getValue() - 1));

        } else if (r[3] === 'Female') {
          Logger.log("Processing Senior Female entry");
          Logger.log(`Gender: ${r[3]}`);
          Logger.log(`Age: ${r[2]}`);
          //Total Senior J8
        var totalSenior = totalSeniorRange.getValue() - 1;
        totalSeniorRange.setValue(parseInt(totalSenior));
        //Total Senior X13
        var totalSeniorMonthly = totalSeniorMonthlyRange.getValue() - 1;
        totalSeniorMonthlyRange.setValue(parseInt(totalSeniorMonthly));
        //SorterSheet
        var seniorFemaleRange = sorterSheet.getRange(sheetRangesSenior[tripNumber].female);
        var seniorFemaleDecrement = seniorFemaleRange.getValue() - 1;
        seniorFemaleRange.setValue(parseInt(seniorFemaleDecrement));
        var infosheetSenior = infoSheet.getRange("D10")
        infosheetSenior.setValue(parseInt(infosheetSenior.getValue() - 1));
        }
      }

      if (r[4] === 'Yes') {
        if (r[3] === 'Male') {
          Logger.log("Processing PWD Male entry");
          Logger.log(`Gender: ${r[3]}`);
          Logger.log(`PWD: ${r[4]}`);
          //Total PWD J8
        var totalPWD = totalPWDRange.getValue() - 1;
        totalPWDRange.setValue(parseInt(totalPWD));
        //total PWD X12
        var totalPWDMonthly = totalPWDMonthlyRange.getValue() - 1;
        totalPWDMonthlyRange.setValue(parseInt(totalPWDMonthly))
        //SorterSheet
        var pwdMaleRange = sorterSheet.getRange(sheetRangesPWD[tripNumber].male);
        var pwdMaleDecrement = pwdMaleRange.getValue() - 1;
        pwdMaleRange.setValue(parseInt(pwdMaleDecrement));
        var infosheetPWD = infoSheet.getRange("D9")
        infosheetPWD.setValue(parseInt(infosheetPWD.getValue() - 1));

        } else if (r[3] === 'Female') {
          Logger.log("Processing PWD Female entry");
          Logger.log(`Gender: ${r[3]}`);
          Logger.log(`PWD: ${r[4]}`);
          //Total PWD J8
        var totalPWD = totalPWDRange.getValue() - 1;
        totalPWDRange.setValue(parseInt(totalPWD));
        //total PWD X12
        var totalPWDMonthly = totalPWDMonthlyRange.getValue() - 1;
        totalPWDMonthlyRange.setValue(parseInt(totalPWDMonthly))
        //SorterSheet
        var pwdFemaleRange = sorterSheet.getRange(sheetRangesPWD[tripNumber].female);
        var pwdFemaleDecrement = pwdFemaleRange.getValue() - 1;
        pwdFemaleRange.setValue(parseInt(pwdFemaleDecrement));
        var infosheetPWD = infoSheet.getRange("D9")
        infosheetPWD.setValue(parseInt(infosheetPWD.getValue() - 1));
        }
      }
    }
  return o;
  }, { v: [], cells: [], rbRanges: [], rbRows: [] });
   if (v.length == 0) {
    Logger.log('No entries to process.');
    return;
  }

    // place newTriggerCancel to row 20
    cells.forEach((cell, index) => {
      var row = parseInt(cell.substring(1));
      sheetNameCBS1.getRange(row, 6).setValue(newTriggerCancel);
      sheetNameCBS1.getRange(row, 1, 1, sheetNameCBS1.getLastColumn()).setBackgroundColor(red);
    });
    Utilities.sleep(500)
    setSortedSheets();
    Logger.log('moveToCancelled() completed.');
}

function setSortedSheets(){
  // Update the sheets using the updated data
  var overallValues = sorterSheet.getRange("A1:AD1").getValues();
  var seniorValues = sorterSheet.getRange("A2:AD2").getValues();
  var pwdValues = sorterSheet.getRange("A3:AD3").getValues();
  var overallRange = sheetNameOverall.getRange(rowOffset_sheets + 1, 3, 1, 30);
  var seniorRange = sheetNameSenior.getRange(rowOffset_sheets + 1, 3, 1, 30);
  var pwdRange = sheetNamePWD.getRange(rowOffset_sheets + 1, 3, 1, 30);

 // Update the ranges
  overallRange.setValues([overallValues[0]]);
  Utilities.sleep(1000); // Adjust the sleep duration as needed
  Logger.log("Sorting Overall Range");

  seniorRange.setValues([seniorValues[0]]);
  Utilities.sleep(1000); // Adjust the sleep duration as needed
  Logger.log("Sorting Senior Range");

  pwdRange.setValues([pwdValues[0]]);
  Utilities.sleep(1000); // Adjust the sleep duration as needed
  Logger.log("Sorting PWD Range");
}

function timeSort(){
  var date = new Date(); // Replace this with your actual date
  Logger.log(date)
  var timeZone = Session.getScriptTimeZone();
  var format = "a";
  Logger.log(formattedTime)
  var formattedTime = Utilities.formatDate(date, timeZone, format);
  var formattedTimeCBS1 = Utilities.formatDate(date, timeZone, format);
  Logger.log(formattedTimeCBS1)

  if(formattedTimeCBS1 == "AM"){
    var timeAM = counterAM.getValue() + 1;
    counterAM.setValue(parseInt(timeAM));

  } else if (formattedTimeCBS1 == "PM"){
    var timePM = sorterSheet.getRange("A17").getValue() + 1;
    counterPM.setValue(parseInt(timePM));
  }
}

function dailyMonitoringReport(callback) {
var placeholders = {
    '#DATE#': date,
    '#BUSDRIVER#': busDriver.getValue(),
    '#BUSCON#': busConductor.getValue(),
    '#PLATENUM#': plateNumber.getValue(),
    '#NOOFTRIPS#': tripTimeAM.getValue() + tripTimePM.getValue(),
    '#TOTALPASS#': infosheetTotal.getValue(),
    '#MALEPASS#': infosheetMale.getValue(),
    '#FEMPASS#': infosheetFemale.getValue(),
    '#TRIPSAM#': tripTimeAM.getValue(),
    '#TRIPSPM#': tripTimePM.getValue(),
    '#MALEAM#': malePassengersAM.getValue(),
    '#FEMAM#': femalePassengersAM.getValue(),
    '#MALEPM#': malePassengersPM.getValue(),
    '#FEMPM#': femalePassengersPM.getValue(),
    '#TOTALAM#': totalPassengersAM.getValue(),
    '#TOTALPM#': totalPassengersPM.getValue()
  };

  // File is the template file, and you get it by ID
  var copy = dmrFile.makeCopy(date + " CBS1_LibrengSakay_DailyMonitoringReport #" + documentNumberDaily, dmrDestFolder);
  var doc = DocumentApp.openById(copy.getId());
  var body = doc.getBody();
  var googleDocLink = doc.getUrl();
  dailyMonitoringReportLink.setValue(googleDocLink);


  // Loop through the placeholders object and replace text
  for (var placeholder in placeholders) {
    body.replaceText(placeholder, placeholders[placeholder]);
  }

  doc.saveAndClose(); 
  callback()
}

function backToZero(callback) {
  var zeroValues = sorterSheet.getRange("A5:AD7").getValues();
  var zeroRange = sorterSheet.getRange("A1:AD3");
  zeroRange.setValues(zeroValues);
  callback();
}

function resetDailySetTables(callback){
var overallValues = sorterSheet.getRange("A1:AD1").getValues();
var seniorValues = sorterSheet.getRange("A2:AD2").getValues();
var pwdValues = sorterSheet.getRange("A3:AD3").getValues();

var overallRange = sheetNameOverall.getRange(rowOffset_sheets + 1, 3, 1, 30);
var seniorRange = sheetNameSenior.getRange(rowOffset_sheets + 1, 3, 1, 30);
var pwdRange = sheetNamePWD.getRange(rowOffset_sheets + 1, 3, 1, 30);

overallRange.setValues([overallValues[0]]);
seniorRange.setValues([seniorValues[0]]);
pwdRange.setValues([pwdValues[0]]);
Logger.log('Spreadsheet Set Tables Successfully.');
callback();
}

function resetDailyMakeCopy(callback){
var newFile = DriveApp.getFileById("16S4s8HgObqYlgpW1625a5KLPu1sn-SOaFx6ZNoZCru4").makeCopy(date + " CBS1_LibrengSakay_ResetCopy #" + documentNumberDaily, dailyDestFolder);
var spreadsheet = SpreadsheetApp.openById(newFile.getId());
var googleDocLink = spreadsheet.getUrl();
dailyResetCopyLink.setValue(googleDocLink);
Logger.log('Spreadsheet Copied Successfully.');
callback();
}

function resetDailySetValues(callback){
  var addRowValue = sorterSheet.getRange('A9').getValue() + 1;
  var addRowMonthlyValue = sorterSheet.getRange('A11').getValue() + 1;
  sorterSheet.getRange('A9').setValue(addRowValue);
  sorterSheet.getRange('A11').setValue(addRowMonthlyValue);
  tripCountSheet.getRange("B2").setValue(1);
  documentCounterDaily.setValue(documentCounterDaily.getValue() + 1);
  infoSheet.getRange("D6").setValue(parseInt(1));
  [counterAM, counterPM].forEach(counter => counter.setValue(0));
  ["D9", "D10", "D11", "D12", "D13", "D14", "D15", "D16", "D17", "D18"].forEach(range => infoSheet.getRange(range).setValue(0));
  Logger.log('Spreadsheet Set Values Successfully.');
  callback();
}

function resetDailyClearFormat(callback){
sheetNameCBS1.getRange('A2:F2000').clearFormat().clearContent();
sorterSheet.getRange('A1:AD3').clearContent();
Logger.log('Spreadsheet Cleared Format.');
callback();
}

function resetDailyGenerateReports() {
        resetDailyMakeCopy(function() {
          resetDailySetValues(function() {
            updateMainMonitoring(function() {
              resetDailyClearFormat(function() {
                backToZero(function() {
                  Logger.log('resetDailyGenerateReports() Succesful.');
          });
        });
      });
    });
  });
}

function resetMonthSetTables(callback){
var overallValues = sorterSheet.getRange("A1:AD1").getValues();
var seniorValues = sorterSheet.getRange("A2:AD2").getValues();
var pwdValues = sorterSheet.getRange("A3:AD3").getValues();

var overallRange = sheetNameOverall.getRange(rowOffset_sheets + 1, 3, 1, 30);
var seniorRange = sheetNameSenior.getRange(rowOffset_sheets + 1, 3, 1, 30);
var pwdRange = sheetNamePWD.getRange(rowOffset_sheets + 1, 3, 1, 30);

overallRange.setValues([overallValues[0]]);
seniorRange.setValues([seniorValues[0]]);
pwdRange.setValues([pwdValues[0]]);
Logger.log('Spreadsheet Set Tables Successfully.');
callback();
}

function resetMonthMakeCopy(callback){
  var newFile = DriveApp.getFileById("16S4s8HgObqYlgpW1625a5KLPu1sn-SOaFx6ZNoZCru4").makeCopy(date + " CBS1_LibrengSakay_MonthlyReport #" + documentNumberMonthly, monthlyDestFolder);
  var spreadsheet = SpreadsheetApp.openById(newFile.getId());
  var googleDocLink = spreadsheet.getUrl();
  dailyResetCopyLink.setValue(googleDocLink);
  Logger.log('Spreadsheet Copied Successfully.');
  callback();
}

function resetMonthSetValues(callback){
totalPassengerRange.setValue(parseInt(0));
  totalSeniorRange.setValue(parseInt(0));
  totalPWDRange.setValue(parseInt(0));
  totalMaleMonthlyRange.setValue(parseInt(0));
  totalFemaleMonthlyRange.setValue(parseInt(0));
  totalSeniorMonthlyRange.setValue(parseInt(0));
  totalPWDMonthlyRange.setValue(parseInt(0));
  totalPassengerMonthlyRange.setValue(parseInt(0));
  infoSheet.getRange("A2").setValue("LAST RECORDED DATE")
  grandTotalTripsMonthly.setValue(parseInt(0));
  grandTotalTripsRange.setValue(parseInt(0));
  [counterAM, counterPM].forEach(counter => counter.setValue(0));
  ["D9", "D10", "D11", "D12", "D13", "D14", "D15", "D16", "D17", "D18"].forEach(range => infoSheet.getRange(range).setValue(0));

  sorterSheet.getRange('A9').setValue(11);
  sorterSheet.getRange('A11').setValue(17);
  infoSheet.getRange("D6").setValue(parseInt(1));
  infoSheet.getRange("D17").setValue(parseInt(0));
  infoSheet.getRange("D26").setValue(parseInt(0));

  tripCountSheet.getRange("B2").setValue(1);

  documentCounterDaily.setValue(documentCounterDaily.getValue() + 1);
  documentCounterMonthly.setValue(documentCounterMonthly.getValue() + 1);
  Logger.log('Spreadsheet Values set.');
  callback();
}

function resetMonthClearFormat(callback){
sheetNameCBS1.getRange('A2:F2000').clearFormat().clearContent();
sheetNameMonthly.getRange('B17:AA47').clearContent();
sheetNameOverall.getRange('B12:AH42').clearContent();
sheetNamePWD.getRange('B12:AH42').clearContent();
sheetNameSenior.getRange('B12:AH42').clearContent();
Logger.log('Spreadsheet Cleared Format.');
callback();
}

function updateTemplatesMonthly(){
  const today = new Date();
  const formattedDate = Utilities.formatDate(today, "UTC+8", "MMMM yyyy");
  var republicOfthePH = "Republic of the Philippines"
  var officeOfthePresident = "Office of the Vice President"
  var ovpLSDTLSheet = "OVP LIBRENG SAKAY DAILY TRIP LOG SHEET"
  var cbsSatelliteOffice = "CEBU, BOHOL & SIQUIJOR (CBS) SATELLITE OFFICE"
  var monthOf = "FOR THE MONTH OF " + formattedDate
  var finaltextMonthly = ("\n \n \n \n \n \n" + republicOfthePH + "\n" + officeOfthePresident + "\n \n" + ovpLSDTLSheet + "\n \n " + cbsSatelliteOffice + "\n" + monthOf)
  var finaltextSheets = ("\n \n \n \n \n" + republicOfthePH + "\n" + officeOfthePresident + "\n \n" + ovpLSDTLSheet + "\n \n " + cbsSatelliteOffice + "\n" + monthOf)

  
  ovpTemplateMonthlySheet.setValue(finaltextMonthly).setFontWeight("bold");
   ovpTemplateOverallSheet.setValue(finaltextSheets).setFontWeight("bold");
    ovpTemplateSeniorSheet.setValue(finaltextSheets).setFontWeight("bold");
     ovpTemplatePWDSheet.setValue(finaltextSheets).setFontWeight("bold");
}

function resetMonthGenerateReports(){
        resetMonthMakeCopy(function() {
          resetMonthSetValues(function() {
            updateMainMonitoring(function() {
                resetMonthClearFormat(function() {
                  backToZero(function() {
                    updateTemplatesMonthly(function() {
                    Logger.log('resetMonthGenerateReports() Successful.');
            });
          });
        });
      });
    });
  });
 }

function updateMainMonitoring(callback){
var dailyFileLink = dailyResetCopyLink.getValue();
var dailyMonitoringLink = dailyMonitoringReportLink.getValue();
var route = infoSheet.getRange("D5").getValue();
var plateNumberCBS1 = plateNumber.getValue();
var propertyKey = '_lastNum';
var lastNum = PropertiesService.getScriptProperties().getProperty(propertyKey) || 0;
lastNum = parseInt(lastNum) + 1;
PropertiesService.getScriptProperties().setProperty(propertyKey, lastNum.toString());
var controlNumber = "OVPLS2024_CBS1_RECORD" + lastNum;
    Logger.log(controlNumber);
ncrMainMonitoringSpreadsheet.getRange(lastrow_mainMonitoring + 1, 1, 1, 6).setValues([[date, controlNumber, plateNumberCBS1, route, dailyFileLink, dailyMonitoringLink]]);
Logger.log('Monitoring Updated.');
callback()
}