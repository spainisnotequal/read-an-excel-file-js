const XLSX = require("xlsx");

/* ------------------------------ */
/* First woorkbook (xlsx version) */
/* ------------------------------ */

const fooWorkbook = XLSX.readFile("./foo.xlsx");

// log every row in every sheet of the workook
fooWorkbook.SheetNames.map((sheet) => {
  console.log("Name of the sheet:", sheet);
  XLSX.utils
    .sheet_to_json(fooWorkbook.Sheets[sheet])
    .map((row) => console.log(row));
});

// Process each sheet individually
const fooSheetNameList = fooWorkbook.SheetNames; // ordered list of the sheets in the fooWorkbook

//first sheet
const dataFirstFooSheet = XLSX.utils.sheet_to_json(
  fooWorkbook.Sheets[fooSheetNameList[0]]
);
console.log("Data from the first sheet");
dataFirstFooSheet.map((row) => console.log(row));

// second sheet
const dataSecondFooSheet = XLSX.utils.sheet_to_json(
  fooWorkbook.Sheets[fooSheetNameList[1]]
);
console.log("Data from the second sheet");
dataSecondFooSheet.map((row) => console.log(row));

/* ------------------------------ */
/* Second woorkbook (xls version) */
/* ------------------------------ */

const barWorkbook = XLSX.readFile("./bar.xls");

// log every row in every sheet of the workook
barWorkbook.SheetNames.map((sheet) => {
  console.log("Name of the sheet:", sheet);
  XLSX.utils
    .sheet_to_json(barWorkbook.Sheets[sheet])
    .map((row) => console.log(row));
});

// Process each sheet individually
const barSheetNameList = barWorkbook.SheetNames; // ordered list of the sheets in the barWorkbook

//first sheet
const dataFirstBarSheet = XLSX.utils.sheet_to_json(
  barWorkbook.Sheets[barSheetNameList[0]]
);
console.log("Data from the first sheet");
dataFirstBarSheet.map((row) => console.log(row));

// second sheet
const dataSecondBarSheet = XLSX.utils.sheet_to_json(
  barWorkbook.Sheets[barSheetNameList[1]]
);
console.log("Data from the second sheet");
dataSecondBarSheet.map((row) => console.log(row));
