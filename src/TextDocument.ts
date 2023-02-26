import { OfficeDocument } from "./OfficeDocument"

export class TextDocument extends OfficeDocument {
    constructor() {
        super()
        this.type = "text"
    }
}