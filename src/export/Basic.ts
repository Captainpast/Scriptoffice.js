import { SpreadsheetDocument } from "../SpreadsheetDocument";

export interface SpreadsheetCsvOptions {
    separator: string,
    useQuotes: boolean | string,
    sheet: string | number
}

function escapeQuotes(value: string, quote: string): string {
    if (quote) {
        let reg = new RegExp(quote, "g")
        return value.replace(reg, quote+quote)
    } else {
        return value;
    }
}

export function spreadsheetToCSV(doc: SpreadsheetDocument, options?: SpreadsheetCsvOptions): string {
    let sheet = doc.activeSheet;
    if (options?.sheet) {
        sheet = doc.getSheet(options?.sheet)
    }
    let separator = options?.separator || ";";
    let quotes = "\"";
    if (options?.useQuotes == false) {
        quotes = "";
    }
    if (typeof(options?.useQuotes) == "string") {
        quotes = options?.useQuotes
    }

    let dataArray = sheet.convertToArray()

    let result = "";
    for (const row of dataArray) {
        for (let i = 0; i < row.length; i++) {
            const data = row[i];
            let value = data?.toString() || "";
            if (data instanceof Date) {
                value = data.toISOString();
            }
            result += quotes+ escapeQuotes(value, quotes) +quotes;

            if ((i+1) < row.length) {
                result += separator;
            }
        }
        result += "\n";
    }
    return result
}