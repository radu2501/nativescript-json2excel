"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var WorkbookSettings = jxl.WorkbookSettings;
var WritableWorkbook = ro.nicoara.radu.jexcelwrapper.WritableWorkbookWrapper;
var WritableFont = jxl.write.WritableFont;
var lodash_1 = require("lodash");
var util = require("./util");
function json2XlsFile(json, path, options) {
    var file = new java.io.File(path);
    var wbSettings = new WorkbookSettings();
    wbSettings.setLocale(new java.util.Locale('en', 'EN'));
    var workbook = WritableWorkbook.createWorkbook(file, wbSettings);
    var sheet = workbook.createSheet(options.sheetName, 0);
    options.fields.forEach(function (field, index) {
        util.addLabel(sheet, index, 0, field.title || field.field);
    });
    json.forEach(function (row, index) {
        options.fields.forEach(function (field, fieldIndex) {
            if (field.type === 'number') {
                return util.addNumber(sheet, fieldIndex, index + 1, lodash_1.get(row, field.field));
            }
            util.addLabel(sheet, fieldIndex, index + 1, ('' + lodash_1.get(row, field.field)));
        });
    });
    if (options.autoSize) {
        sheet.setAutoSize(true);
    }
    if (options.width && options.width > 0) {
        sheet.setColumnsWidth(options.width);
    }
    workbook.write();
    workbook.close();
}
exports.json2XlsFile = json2XlsFile;
