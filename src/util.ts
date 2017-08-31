
declare var jxl;
declare var ro;

let WritableCellFormat = ro.nicoara.radu.jexcelwrapper.WritableCellFormatWrapper;
let WritableFont = jxl.write.WritableFont;
let LabelFactory = ro.nicoara.radu.jexcelwrapper.LabelFactory;
let cellFont = new WritableFont(WritableFont.ARIAL, 10);
let format = new jxl.write.WritableCellFormat(cellFont);
format.setWrap(true);

export function addLabel(sheet: any, column: number, row: number, text: string) {
    let label = LabelFactory.createLabel(column, row, text);
    label.setCellFormat(format);
    sheet.addLabelCell(label);
}

export function addNumber(sheet: any, column: number, row: number, int: number) {
    let nr = LabelFactory.createNumber(column, row, int);
    nr.setCellFormat(format);
    sheet.addNumberCell(nr);
}

