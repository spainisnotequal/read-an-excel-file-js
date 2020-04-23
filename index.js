const XLSX = require("xlsx");

/* ------------------------------ */
/* First woorkbook (xlsx version) */
/* ------------------------------ */

const workbook = XLSX.readFile("./foo.xlsx");
const sheetNameList = workbook.SheetNames; // ordered list of the sheets in the workbook

//first sheet
const dataFirstSheet = XLSX.utils.sheet_to_json(
  workbook.Sheets[sheetNameList[0]]
);
console.log("Data from the first sheet");
dataFirstSheet.map((row) => console.log(row));

// second sheet
const dataSecondSheet = XLSX.utils.sheet_to_json(
  workbook.Sheets[sheetNameList[1]]
);
console.log("Data from the second sheet");
dataSecondSheet.map((row) => console.log(row));

/* ------------------------------ */
/* Second woorkbook (xls version) */
/* ------------------------------ */

const workbook2 = XLSX.readFile("./bar.xls");
const sheetNameList2 = workbook.SheetNames; // ordered list of the sheets in the workbook

//first sheet
const dataFirstSheet2 = XLSX.utils.sheet_to_json(
  workbook2.Sheets[sheetNameList2[0]]
);
console.log("Data from the first sheet");
dataFirstSheet2.map((row) => console.log(row));

// second sheet
const dataSecondSheet2 = XLSX.utils.sheet_to_json(
  workbook2.Sheets[sheetNameList2[1]]
);
console.log("Data from the second sheet");
dataSecondSheet2.map((row) => console.log(row));
