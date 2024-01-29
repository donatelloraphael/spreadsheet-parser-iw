import * as process from "process";
import * as path from "path";
import * as ExcelJS from "exceljs";

if (process.argv.length < 3) {
  console.error("Usage: node parser.js [file]");
  process.exit(1);
}

const filename = process.argv[2];

if (path.extname(filename) !== ".xlsx") {
  console.error("Error: The file is not an .xlsx file");
  process.exit(1);
}

// Could've used a streaming API but this is simpler and there are some
// lack of functionality with the streaming API for this library. Could've used another library,
// but given the time constraint, I decided to go with this.
let workbook = new ExcelJS.Workbook();
workbook.xlsx
  .readFile(filename)
  .then(() => {
    workbook.eachSheet((worksheet, sheetId) => {
      let data: Record<string, any>[] = [];
      let schema: Record<string, any> = {};
      let template: Record<string, any> = {};
      let columns: string[] = [];
      worksheet.eachRow((row, rowNumber) => {
        template = JSON.parse(JSON.stringify(schema));
        if (rowNumber === 1) {
          row.eachCell((cell, colNumber) => {
            columns.push(cell.text);
            const parts = cell.text.split("__");
            appendPartsToSchema(schema, parts);
          });
        } else {
          row.eachCell((cell, colNumber) => {
            const text = cell.text;
            const parts = columns[colNumber - 1].split("__");
            appendValuesToSchema(template, parts, text);
          });
          // output each row value to stdout stream without using console.log
          // to avoid extra newline
          process.stdout.write(JSON.stringify(template));

        }
      });
    });
  })
  .catch((err) => {
    console.error("Error reading file: ", err);
  });

function appendPartsToSchema(schema: Record<string, any>, parts: string[]) {
  if (!parts.length) {
    return;
  }
  if (schema[parts[0]] === undefined) {
    schema[parts[0]] = null;
  }
  if (parts[1]) {
    if (schema[parts[0]] === null) {
      schema[parts[0]] = {};
    } else {
      schema[parts[0]][parts[1]] = null;
    }
    appendPartsToSchema(schema[parts[0]], parts.slice(1));
  }
}

function appendValuesToSchema(schema: Record<string, any>, parts: string[], text: string) {
  if (!parts.length) {
    return;
  }
  if (parts[1]) {
    appendValuesToSchema(schema[parts[0]], parts.slice(1), text);
  } else {
    const splitText = text.split(';').map((text) => text?.trim()?.toLowerCase());
    schema[parts[0]] = splitText.length > 1 ? splitText : (splitText[0] || null);
  }
}
