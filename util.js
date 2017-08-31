"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var WritableCellFormat = ro.nicoara.radu.jexcelwrapper.WritableCellFormatWrapper;
var WritableFont = jxl.write.WritableFont;
var LabelFactory = ro.nicoara.radu.jexcelwrapper.LabelFactory;
var cellFont = new WritableFont(WritableFont.ARIAL, 10);
var format = new jxl.write.WritableCellFormat(cellFont);
format.setWrap(true);
function addLabel(sheet, column, row, text) {
    var label = LabelFactory.createLabel(column, row, text);
    label.setCellFormat(format);
    sheet.addLabelCell(label);
}
exports.addLabel = addLabel;
function addNumber(sheet, column, row, int) {
    var nr = LabelFactory.createNumber(column, row, int);
    nr.setCellFormat(format);
    sheet.addNumberCell(nr);
}
exports.addNumber = addNumber;
