const nextChar = (char) =>
  char ? String.fromCharCode(char.charCodeAt(0) + 1) : "A";

const nextCol = (str) =>
  str.replace(/([^Z]?)(Z*)$/, (_, a, z) => nextChar(a) + z.replace(/Z/g, "A"));

const fieldNamesToFieldObject = (fieldNames) =>
  fieldNames.reduce((result, element) => {
    result[element] = null;
    return result;
  }, {});

const getColumnsRange = (cellsRange) => cellsRange.match(/[a-zA-Z]+/g);
const getRowsRange = (cellsRange) =>
  cellsRange.match(/[0-9]+/g).map((item) => parseInt(item)); // match returns strings but I want integers

const getFields = (sheet) => {
  const cellsRange = sheet["!ref"];
  const [firstColumn, lastColumn] = getColumnsRange(cellsRange);
  const [firstRow, lastRow] = getRowsRange(cellsRange);
  console.log(firstColumn, lastColumn, firstRow, lastRow);

  //let fields = [];

  for (let col = firstColumn; col !== nextCol(lastColumn); col = nextCol(col)) {
    for (let row = firstRow + 1; row <= lastRow; row++) {
      const cellId = col + row;
      console.log(cellId);
    }
  }
};
