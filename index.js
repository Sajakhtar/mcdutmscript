// SheetJS
// To install: $ npm install xlsx
// To run: node index.js
// https://www.npmjs.com/package/xlsx
// https://github.com/SheetJS/js-xlsx
// https://github.com/SheetJS/js-xlsx/tree/1eb1ec985a640b71c5b5bbe006e240f45cf239ab/demos/server
// https://github.com/SheetJS/js-xlsx/blob/1eb1ec985a640b71c5b5bbe006e240f45cf239ab/demos/server/express.js
// https://github.com/SheetJS/js-xlsx/blob/1eb1ec985a640b71c5b5bbe006e240f45cf239ab/demos/server/node.js
// https://github.com/SheetJS/js-xlsx/blob/master/tests/write.js
// https://thewebspark.com/2018/05/13/how-to-access-excel-sheets-data-using-node-js/

const path = require("path");
const fs = require("fs");

if (typeof require !== "undefined") XLSX = require("xlsx");

const readline = require("readline-sync");

// Campaign name user input
// https://teamtreehouse.com/community/how-to-get-input-in-the-console-in-nodejs
var campaign = readline.question("What is the campaign name? ");
//console.log(campaign);

var wb = XLSX.readFile("testtest.xlsx");
//console.log(testWorkbook);

var wsName = wb.SheetNames[0];
var ws = wb.Sheets[wsName];
//console.log(ws1);

// READ: XLSX.utils.sheet_to_json
// array of objects
// WRITE: XLSX.utils.json_to_sheet
var wsJSON = XLSX.utils.sheet_to_json(ws);
//console.log(wsJSON);

let dataIn = wsJSON.filter(x => x["Placement"] != "");
//console.log(dataIn);

let dataOut = [...dataIn];
//console.log(dataOut);

dataOut.map(x => {
  // UTMs
  //utm_source=banner&utm_medium={publisher}&utm_campaign={User input}&utm_content={size_creative type_market_language}

  // Remove key and ' (AE)'
  let publisher = x["Placement"]
    .split("_")[5]
    .split("~")[1]
    .split(" (")[0];

  let size = x["Placement"].split("_")[7].split("~")[1];

  let creativeType = x["Placement"].toLowerCase().includes("video")
    ? "video"
    : x["Placement"].toLowerCase().includes("1x1")
    ? "tracker"
    : "html5";

  let market = x["Placement"].split("_")[11].split("~")[1];

  let lang = x["Placement"]
    .split("_")[10]
    .split("~")[1]
    .slice(0, 2);

  let content = `${size}_${creativeType}_${market}_${lang}`;

  // replace special characters REGEX
  // var regExpr = /[^a-zA-Z0-9-. ]/g
  // .replace(regExpr, "")
  // .replace(/[^a-zA-Z0-9]/g, "")

  let utm = `utm_source=banner&utm_medium=${publisher
    .toLowerCase()
    .replace(" - ", "")
    .replace(" ", "")}&utm_campaign=${campaign
    .toLowerCase()
    .replace(" ", "")}&utm_content=${content.toLowerCase()}`;

  //check URL includes "?"
  let queryParam = x["URL 1"].includes("?") ? "&" : "?";

  x["URL 1"] = x["URL 1"] + queryParam + utm;
});
//console.log(dataOut);

// Get column header
let columnHeaders = Object.keys(dataOut[0]);
//console.log(columnHeaders);

// ******************************************

var workbookOut = XLSX.utils.book_new();
var wsOutName = wsName;
var wsOut = XLSX.utils.json_to_sheet(dataOut, { header: columnHeaders });

XLSX.utils.book_append_sheet(workbookOut, wsOut, wsOutName);

/* output format determined by filename */
XLSX.writeFile(workbookOut, "./testtestOut.xlsx");

//const folderPath = "C:\\Users\\Sajakhta\\Desktop";
//const folderPath = "C:\\Users\\Sajakhta\\Documents";
//const folderPath = "C:\\Users\\Sajakhta\\Downloads";
//const fileName = "\\SheetJS2.xlsx";
//const filePath = folderPath + fileName;
//console.log(filePath);
//XLSX.writeFile(workbook, filePath);
