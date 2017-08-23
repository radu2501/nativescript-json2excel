export declare function json2XlsxFile(json: any[], path: string, options: JsonToXlsxOptions): void;
export interface JsonToXlsxOptions {
    fields: {
        title?: string;
        field: string;
    }[];
    sheetName: string;
}
