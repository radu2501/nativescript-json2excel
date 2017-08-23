declare var java;
declare var org;

let Cell = org.apache.poi.ss.usermodel.Cell;
let CreationHelper = org.apache.poi.ss.usermodel.CreationHelper;
let Row = org.apache.poi.ss.usermodel.Row;
let Sheet = org.apache.poi.ss.usermodel.Sheet;
let Workbook = org.apache.poi.ss.usermodel.Workbook;
let XSSFWorkbook = org.apache.poi.xssf.usermodel.XSSFWorkbook;

import { get } from 'lodash';

export function json2XlsxFile(json: any[], path: string, options: JsonToXlsxOptions): void {
    let wb = new XSSFWorkbook();
    let helper = wb.getCreationHelper();

    let sheet = wb.createSheet(options.sheetName);

    let headerRow = sheet.createRow(0);
    options.fields.forEach((field, index) => {
        headerRow.createCell(index).setCellValue(helper.createRichTextString(field.title || field.field));
    });
    json.forEach((row, index) => {
        let xRow = sheet.createRow(index + 1);
        options.fields.forEach((field, fieldIndex) => {
            xRow.createCell(fieldIndex).setCellValue(get(row, field.field));
        });
    });

    let fos = new java.io.FileOutputStream(path);
    wb.write(fos);
}

export interface JsonToXlsxOptions {
    fields: { title?: string, field: string }[],
    sheetName: string
}