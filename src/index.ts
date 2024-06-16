//@ts-nocheck
import { OfficeDocumentTypes } from "./OfficeDocument"
import { SpreadsheetDocument } from "./SpreadsheetDocument"
import { TextDocument } from "./TextDocument"

function create<T extends keyof OfficeDocumentTypes>(options: T | { type: T }): OfficeDocumentTypes[T] {
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
function load(src: string) {
    throw "not implemented\n"
}

export const OfficeDocument =  { create, load }