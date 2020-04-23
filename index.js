const XLSX = require("xlsx");

// read the first sheet of the "foo.xlsx" file and get its fields
const FILE = "./foo.xlsx";
const workbook = XLSX.readFile(FILE);

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
const fields = Object.keys(sheetData[0]);
console.log(fields);

// get the data in A2
console.log("type:", sheet.A2.t); //=> "n" means number
console.log("raw value:", sheet.A2.v);
console.log("formatted text:", sheet.A2.w);

// get the data in B2
console.log("type:", sheet.B2.t); //=> "s" means text
console.log("raw value:", sheet.B2.v);
console.log("formatted text:", sheet.B2.w);

// get the data in C2
console.log("type:", sheet.C2.t); //=> "n" means number
console.log("raw value:", sheet.C2.v);
console.log("formatted text:", sheet.C2.w);
