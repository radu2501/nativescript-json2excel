export declare function json2XlsFile(json: any[], path: string, options: JsonToXlsxOptions): void;
export interface JsonToXlsxOptions {
    fields: {
        title?: string;
        field: string;
        type?: string;
    }[];
    sheetName: string;
    autoSize?: boolean;
    width?: number;
}
