const excelToJson = require("convert-excel-to-json");
const fs = require("fs");

/**
 * @description
 * In case all objects in an array of objects doesn't have some missing keys then this function will append those missing keys in all other objects and make it equal
 */
function handleMakeEqual(arr) {
  const keys = arr.reduce(
    (acc, curr) => (Object.keys(curr).forEach((key) => acc.add(key)), acc),
    new Set()
  );

  const output = arr.map((item) =>
    [...keys].reduce((acc, key) => ((acc[key] = item[key] ?? ""), acc), {})
  );
  return output;
}

function filter(data = []) {
  const arr = handleMakeEqual(data);
  const header = arr[0];
  const body = arr.slice(1);
  const raw = body.map((item) => {
    const instance = {};
    for (const key in item) {
      let newInstanceKey = header[key];
      if (newInstanceKey in instance) {
        // already another key exists in instance
        const reg = new RegExp(newInstanceKey, "gi");
        const noOfOccurence = Object.keys(instance).filter((item) =>
          reg.test(item)
        ).length;
        newInstanceKey = `${newInstanceKey}-${noOfOccurence}`;
        instance[newInstanceKey] = item[key];
      }

      if (
        /Invoice Date/gi.test(newInstanceKey) ||
        /Cancellation Date/gi.test(newInstanceKey)
      ) {
        // date validation
        if (!item[key]) {
          instance[newInstanceKey] = "";
        } else if (new Date(item[key]) == "Invalid Date") {
          instance[newInstanceKey] = item[key];
        } else {
          instance[newInstanceKey] = new Date(item[key]).toLocaleDateString();
        }
      } else if (item[key] || item[key] === 0) {
        // item[key] has a zero as value or some value
        instance[newInstanceKey] = item[key];
      } else {
        // item[key] is null or undefined
        instance[newInstanceKey] = "";
      }
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
