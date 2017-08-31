declare var java;
declare var jxl;
declare var ro;

let WorkbookSettings = jxl.WorkbookSettings;
let WritableWorkbook = ro.nicoara.radu.jexcelwrapper.WritableWorkbookWrapper;
let WritableFont = jxl.write.WritableFont;

import { get } from 'lodash';
import * as util from './util';

export function json2XlsFile(json: any[], path: string, options: JsonToXlsxOptions): void {
    let file = new java.io.File(path);
    let wbSettings = new WorkbookSettings();

    wbSettings.setLocale(new java.util.Locale('en', 'EN'));

    let workbook = WritableWorkbook.createWorkbook(file, wbSettings);
    let sheet = workbook.createSheet(options.sheetName, 0);


    options.fields.forEach((field, index) => {
        util.addLabel(sheet, index, 0, field.title || field.field);
    });
    json.forEach((row, index) => {
        options.fields.forEach((field, fieldIndex) => {
            if (field.type === 'number') {
                return util.addNumber(sheet, fieldIndex, index + 1, get(row, field.field) as number);
            }
            util.addLabel(sheet, fieldIndex, index + 1, ('' + get(row, field.field)) as string);
        });
    });

    workbook.write();
    workbook.close();
}

export interface JsonToXlsxOptions {
    fields: { title?: string, field: string, type?: string }[];
    sheetName: string;
}



