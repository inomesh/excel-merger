const excelToJson = require("convert-excel-to-json");
const fs = require("fs");

function filter(data = []) {
  const header = data[0];
  const body = data.slice(1);
  const raw = body.map((item) => {
    const instance = {};
    for (const key in item) {
      instance[header[key]] =
        item[key] === null && item[key] === undefined ? "" : item[key];
    }
    return instance;
  });
  return raw;
}

function ExcToJSON(excelFilePath) {
  const obj = {};
  const result = excelToJson({
    source: fs.readFileSync(excelFilePath),
  });
  for (const key in result) {
    const json = filter(result[key]);
    obj[key] = json;
  }
  return obj;
}

module.exports = { ExcToJSON };
