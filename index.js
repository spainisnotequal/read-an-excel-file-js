const XLSX = require("xlsx");

/* ------------------------------ */
/* First woorkbook (xlsx version) */
/* ------------------------------ */

const fooWorkbook = XLSX.readFile("./foo.xlsx");

// log every row in every sheet of the workook
fooWorkbook.SheetNames.map((sheetName) => {
  console.log("Name of the sheet:", sheetName);

  const sheet = XLSX.utils.sheet_to_json(fooWorkbook.Sheets[sheetName]);

  if (sheet[0]) console.log("Fields of the sheet:", Object.keys(sheet[0]));
  else console.log("The sheet is empty.");

  // Log each row of the sheet
  sheet.map((row) => console.log(row));
});

/* ------------------------------ */
/* Second woorkbook (xls version) */
/* ------------------------------ */

const barWorkbook = XLSX.readFile("./bar.xls");

// log every row in every sheet of the workook
barWorkbook.SheetNames.map((sheetName) => {
  console.log("Name of the sheet:", sheetName);

  const sheet = XLSX.utils.sheet_to_json(barWorkbook.Sheets[sheetName]);

  if (sheet[0]) console.log("Fields of the sheet:", Object.keys(sheet[0]));
  else console.log("The sheet is empty.");

  // Log each row of the sheet
  sheet.map((row) => console.log(row));
});
