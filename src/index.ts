//@ts-nocheck
import { OfficeDocumentTypes } from "./OfficeDocument"
import { SpreadsheetDocument } from "./SpreadsheetDocument"
import { TextDocument } from "./TextDocument"

export function create<T extends keyof OfficeDocumentTypes>(options: T | { type: T }): OfficeDocumentTypes[T] {
    if (typeof(options) == "string") {
        options = { type: options }
    }
    switch (options.type) {
        case "spreadsheet":
            return new SpreadsheetDocument()
        case "text":
            return new TextDocument()
        default:
            throw `${options.type} is not a supported document type`
    }
}

/**@deprecated WIP*/
export function load(src: string) {
    throw "not implemented\n"
}
