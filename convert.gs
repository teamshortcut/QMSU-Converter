/*
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Paste your input below this!
INSTRUCTIONS:
1. Export your data from QMSU as a CSV.
2. Import to a temporary sheet (File -> Import) - make sure you tick ""Convert text to numbers, dates, and formulas""
3. Copy the entire sheet below the divider
4. Delete any previous results!
Delete any leftover results from below the second divider, if there is anything left (make sure you've copied or saved it elsewhere if you needed it!)
5. From the menu, select Custom QMSU menu -> Process data from a SINGLE academic year
6. Wait for the script to finish running, then copy your results!

Source code can be found here: https://script.google.com/home/projects/1ob21s7Wv2R0cmblTvzWBlnDbQffycW97ssbOkIYE_jV3otbP0OAVr15Z/edit
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
*/

// Extensions -> Apps Script

const sheet = SpreadsheetApp.getActiveSheet()

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu("Custom QMSU menu");
  menu.addItem("Process data from a SINGLE academic year", "processCsv");
}

function isMembership(row) {
  const product = row[0].toLowerCase();
  return product.includes("standard") ||
         product.includes("virtual") ||
         product.includes("associate") ||
         product.includes("membership");
}

// ===== PROCESS DATA FOR SINGLE YEAR =====

// NOTE: Assumes data is from a single (academic) year!
function processCsv() {
  let data = sheet.getDataRange().getValues();
  data.shift(); // remove instructions line
  data.splice(0, 4) // remove first 4 rows, as they only contain metadata
  data.sort((a, b) => a[7] - b[7]);

  const results = {total: 0}

  for (let i = 0; i < data.length; i++) {
    const row = data[i];

    // check is this row represents a membership purchase
    // (rather than event ticket, for example)
    if (!isMembership(row)) {
      continue;
    }

    // reformat full name as "FirstName LastName"
    const nameSplit = row[2].split(",");
    const name = nameSplit[1].slice(1) + " " + nameSplit[0];

    // get student id
    const studentId = row[4].toString();

    // get membership type
    const product = row[0].toLowerCase();
    let membershipType = "";
    if (product.includes("virtual")) {
      membershipType = "Virtual";
    }
    else if (product.includes("associate")) {
      membershipType = "Associate";
    }
    else {
      membershipType = "Standard";
    }

    const date = row[7]; // JS date

    // add record for name, assuming we haven't encountered it before
    // theoretically, we shouldn't encounter a name more than once, assuming everyone has a unique full name
    // however, there is a possibility that a non-membership product slips through, or that two people have the same name
    // it's even possible that someone refunds their membership, causing their name to appear multiple times
    // all of these events will cause the count to be incorrect, but at least won't cause the script to crash 😅
    if (!results.hasOwnProperty(name)) {
      const total = results["total"] + 1
      results["total"] = total
      results[name] = {"studentId": studentId, "date": date, "membershipType": membershipType, "runningTotal": total}
    }
    else {
      Logger.log("DUPLICATE NAME ENCOUNTERED! name=" + name);
    }
  }

  outputResults(results);
}

function outputResults(results) {
  sheet.appendRow(["---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"]);
  sheet.appendRow(["name", "studentId", "date", "runningTotal", "membershipType"])

  delete results["total"] // no longer needed
  for (const [studentName, record] of Object.entries(results)) {
    sheet.appendRow([studentName, record["studentId"], record["date"], record["runningTotal"], record["membershipType"]])
  }
}

// ===== PROCESS COMBINED DATA FOR ALL YEARS =====

// https://stackoverflow.com/a/1184359
function daysInMonth (month, year) {
    return new Date(year, month, 0).getDate();
}

function dateToDayId(date) {
  // NOTE FOR FUTURE COMMITTEES: 2024 (and future years) will need adding to these structures
  const FIRST_SESSION_DATES = {
    2018: new Date(2018, 8, 25),
    2019: new Date(2019, 8, 24),
    2020: new Date(2020, 8, 22),
    2021: new Date(2021, 8, 28),
    2022: new Date(2022, 8, 27),
    2023: new Date(2023, 8, 26) // TODO: This is a guess!
  }
  const FIRST_SESSION_OFFSETS = {
    2018: 2,
    2019: 3,
    2020: 5,
    2021: -1,
    2022: 0,
    2023: 1 // TODO: This is a guess! (based on 26th start date above)
  }

  let month = date.getMonth(); // 0-indexed


  let offsetYear = date.getFullYear();
  let year = 2023 // NOTE FOR FUTURE COMMITTEES: This, and the year below, will need incrementing by 1
  if (month < 7) { // before August
    offsetYear -= 1;
    year = 2024;
  }

  day = date.getDate() + FIRST_SESSION_OFFSETS[offsetYear];
  maxDays = daysInMonth(month, date.getFullYear());
  if (day > maxDays) {
    day = day - maxDays
    month += 1

    if (month > 11) {
      month = 0
      year += 1
    }
  }
  if (day < 1) {
    month -= 1
    day = daysInMonth(month, date.getFullYear()) - day

    if (month < 0) {
      month = 11
      year -= 1
    }
  }

  return new Date(year, month, day)
}

function processCombinedCsv() {
  sheet.getRange("A2:G300").deleteCells(SpreadsheetApp.Dimension.COLUMNS);

  const results = {};

  for (let year = 2018; year <= 2023; year++) { // NOTE FOR FUTURE COMMITTEES: This year will need incrementing by 1
    const yearSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(year);
    const data = yearSheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const date = row[2];

      if (!results.hasOwnProperty(year)) {
        results[year] = {"total": 0}
      }
      results[year]["total"] = results[year]["total"] + 1

      const dateId = dateToDayId(date)

      if (!results[year].hasOwnProperty(dateId)) {
        results[year][dateId] = 0
      }
      results[year][dateId] = results[year][dateId] + 1
    }
  }

  // output
  sheet.appendRow(["date", "2018", "2019", "2020", "2021", "2022", "2023"]) // NOTE FOR FUTURE COMMITTEES: This year will need incrementing by 1
  const totals = {
    2018: 0,
    2019: 0,
    2020: 0,
    2021: 0,
    2022: 0,
    2023: 0,
  }

  let currentDate = new Date(2023, 7, 1); // NOTE FOR FUTURE COMMITTEES: This, and the year below, will need incrementing by 1
  while (currentDate <= new Date(2024, 6, 31)) {
    let = outputLine = false;

    for (const year of Object.keys(results)) {
      if (results[year].hasOwnProperty("total")) {
        delete results[year]["total"]
      }

      if (results[year].hasOwnProperty(currentDate)) {
        outputLine = true
        totals[year] = totals[year] + results[year][currentDate]
        delete results[year][currentDate]
      }
    }

    if (outputLine) {
      sheet.appendRow([currentDate, totals[2018], totals[2019], totals[2020], totals[2021], totals[2022], totals[2023]]); // NOTE FOR FUTURE COMMITTEES: This year will need incrementing by 1
    }

    currentDate.setDate(currentDate.getDate() + 1);
  }
}
