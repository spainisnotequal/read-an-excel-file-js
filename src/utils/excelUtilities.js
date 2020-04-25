const nextChar = (char) =>
  char ? String.fromCharCode(char.charCodeAt(0) + 1) : "A";

const nextCol = (str) =>
  str.replace(/([^Z]?)(Z*)$/, (_, a, z) => nextChar(a) + z.replace(/Z/g, "A"));

const fieldNamesToFieldObject = (fieldNames) =>
  fieldNames.reduce((result, element) => {
    result[element] = null;
    return result;
  }, {});

export default { fieldNamesToFieldObject };
