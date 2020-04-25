const XLSX = require("xlsx");

// read the first sheet of the "foo.xlsx" file and get its fields
const FILE = "./foo.xlsx";
const readOptions = { cellDates: true, cellNF: true, cellText: true };
const workbook = XLSX.readFile(FILE, readOptions);

// log the number of sheets
const numberOfSheets = workbook.Props.Worksheets;
console.log("Number of sheets:", numberOfSheets);

// log the names of the sheets
const sheetsNames = workbook.SheetNames;
console.log("Sheets names:", sheetsNames);

// only consider the first sheet
const sheet = workbook.Sheets[sheetsNames[0]];
console.log(sheet);
const sheetData = XLSX.utils.sheet_to_json(sheet);
console.log(sheetData);
const fieldNames = Object.keys(sheetData[0]);
console.log(fieldNames);

// get the number of rows
const numberOfRows = sheetData.length;
console.log("Number of rows", numberOfRows);

// get the data in A2 (is a Date field)
if (sheet.A2) {
  console.log("type:", sheet.A2.t); //=> "d" means date
  console.log("raw value:", sheet.A2.v);
  console.log("formatted text:", sheet.A2.w);
  console.log("Excel format:", sheet.A2.z);
} else console.log("Empy cell");

// get the data in B2 (is a Concept field)
if (sheet.B2) {
  console.log("type:", sheet.B2.t); //=> "s" means text
  console.log("raw value:", sheet.B2.v);
  console.log("formatted text:", sheet.B2.w);
  console.log("Excel format:", sheet.B2.z);
} else console.log("Empy cell");

// get the data in C2 (is an Amount field)
if (sheet.C2) {
  console.log("type:", sheet.C2.t); //=> "n" means number
  console.log("raw value:", sheet.C2.v);
  console.log("formatted text:", sheet.C2.w);
  console.log("Excel format:", sheet.C2.z);
} else console.log("Empy cell");
