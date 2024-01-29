"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (k !== "default" && Object.prototype.hasOwnProperty.call(mod, k)) __createBinding(result, mod, k);
    __setModuleDefault(result, mod);
    return result;
};
Object.defineProperty(exports, "__esModule", { value: true });
const process = __importStar(require("process"));
const path = __importStar(require("path"));
const ExcelJS = __importStar(require("exceljs"));
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
        let data = [];
        let schema = {};
        let template = {};
        let columns = [];
        worksheet.eachRow((row, rowNumber) => {
            template = JSON.parse(JSON.stringify(schema));
            if (rowNumber === 1) {
                row.eachCell((cell, colNumber) => {
                    columns.push(cell.text);
                    const parts = cell.text.split("__");
                    appendPartsToSchema(schema, parts);
                });
            }
            else {
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
function appendPartsToSchema(schema, parts) {
    if (!parts.length) {
        return;
    }
    if (schema[parts[0]] === undefined) {
        schema[parts[0]] = null;
    }
    if (parts[1]) {
        if (schema[parts[0]] === null) {
            schema[parts[0]] = {};
        }
        else {
            schema[parts[0]][parts[1]] = null;
        }
        appendPartsToSchema(schema[parts[0]], parts.slice(1));
    }
}
function appendValuesToSchema(schema, parts, text) {
    if (!parts.length) {
        return;
    }
    if (parts[1]) {
        appendValuesToSchema(schema[parts[0]], parts.slice(1), text);
    }
    else {
        const splitText = text.split(';').map((text) => { var _a; return (_a = text === null || text === void 0 ? void 0 : text.trim()) === null || _a === void 0 ? void 0 : _a.toLowerCase(); });
        schema[parts[0]] = splitText.length > 1 ? splitText : (splitText[0] || null);
    }
}
//# sourceMappingURL=parser.js.map