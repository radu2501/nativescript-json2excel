"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var Cell = org.apache.poi.ss.usermodel.Cell;
var CreationHelper = org.apache.poi.ss.usermodel.CreationHelper;
var Row = org.apache.poi.ss.usermodel.Row;
var Sheet = org.apache.poi.ss.usermodel.Sheet;
var Workbook = org.apache.poi.ss.usermodel.Workbook;
var XSSFWorkbook = org.apache.poi.xssf.usermodel.XSSFWorkbook;
var lodash_1 = require("lodash");
function json2XlsxFile(json, path, options) {
    var wb = new XSSFWorkbook();
    var helper = wb.getCreationHelper();
    var sheet = wb.createSheet(options.sheetName);
    var headerRow = sheet.createRow(0);
    options.fields.forEach(function (field, index) {
        headerRow.createCell(index).setCellValue(helper.createRichTextString(field.title || field.field));
    });
    json.forEach(function (row, index) {
        var xRow = sheet.createRow(index + 1);
        options.fields.forEach(function (field, fieldIndex) {
            xRow.createCell(fieldIndex).setCellValue(lodash_1.get(row, field.field));
        });
    });
    var fos = new java.io.FileOutputStream(path);
    wb.write(fos);
}
exports.json2XlsxFile = json2XlsxFile;
