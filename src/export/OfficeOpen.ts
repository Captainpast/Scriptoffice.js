import * as JSZip from "jszip";
import { SpreadsheetDocument, SpreadsheetDocumentSheet, getCellName } from "../SpreadsheetDocument";
import { SpreadsheetOptions, dateToString, escapeXML } from "./OpenDocument";


type CSpreadsheetDocument = SpreadsheetDocument & { _strings: string[] }

export async function spreadsheet(doc: CSpreadsheetDocument, options: SpreadsheetOptions): Promise<ArrayBuffer> {
    var zip = new JSZip();
    doc._strings = [];

    zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="xml" ContentType="application/xml" />
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
        <Default Extension="png" ContentType="image/png" />
        <Default Extension="jpeg" ContentType="image/jpeg" />
        <Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
        <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" />
        <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml" />
        <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml" />
        <Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
        <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />
        <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" />
    </Types>`)

    zip.file("docProps/app.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties><Application>ScriptOffice/1.0$Web_UnixLike ${escapeXML(doc.generator)}</Application></Properties>`)
    zip.file("docProps/core.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
        <dc:creator>${escapeXML(doc.initialCreator)}</dc:creator>    
        <dcterms:created xsi:type="dcterms:W3CDTF">${dateToString(doc.creationDate)}</dcterms:created>
        <cp:lastModifiedBy>${escapeXML(doc.creator)}</cp:lastModifiedBy>
        <dcterms:modified xsi:type="dcterms:W3CDTF">${dateToString(doc.date)}</dcterms:modified>
        <dc:title></dc:title>
        <dc:subject></dc:subject>
        <dc:description></dc:description>
    </cp:coreProperties>`)

    zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Target="xl/workbook.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"/>
        <Relationship Id="rId2" Target="docProps/core.xml" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"/>
        <Relationship Id="rId3" Target="docProps/app.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"/>
    </Relationships>`)

    zip.file("xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <sheets>${
            doc.sheets.map((s, i) => `<sheet name="${escapeXML(s.title)}" sheetId="${i+1}" state="visible" r:id="rTableId${i+1}"/>`)
        }</sheets>
    </workbook>`)

    zip.file("xl/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <numFmts count="2">
            <numFmt numFmtId="164" formatCode="General"/>
            <numFmt numFmtId="165" formatCode="dd/mm/yyyy\\ hh:mm"/>
        </numFmts>
        <cellXfs count="2">
            <xf numFmtId="164"></xf>
            <xf numFmtId="165"></xf>
        </cellXfs>
    </styleSheet>`)

    zip.file("xl/_rels/workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Target="styles.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"/>
        <Relationship Id="rId2" Target="sharedStrings.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"/>
        ${
            doc.sheets.map((s, i) => `<Relationship Id="rTableId${i+1}" Target="worksheets/sheet${i+1}.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"/>`)
        }
    </Relationships>`)

    for (let i = 0; i < doc.sheets.length; i++) {
        zip.file(`xl/worksheets/sheet${i+1}.xml`, spreadsheetSheet(doc.sheets[i], doc))
    }

    zip.file("xl/sharedStrings.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="5" uniqueCount="5">
    ${
        doc._strings.map(s => `<si><t xml:space="preserve">${escapeXML(s)}</t></si>`)
    }
    </sst>`)

    var file = await zip.generateAsync({
        type: "arraybuffer",
        platform: "UNIX",
        mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        compression: "DEFLATE",
        compressionOptions: { level: options?.compressionLevel || 9 } })
    return file;
}

const T_1899_12_30 = -2209161600000
const T_HOUR = 3600000
const T_DAY = 86400000

function spreadsheetSheet(sheet: SpreadsheetDocumentSheet, doc: CSpreadsheetDocument): string {
    var tableString = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">`;
    tableString += `<sheetViews><sheetView tabSelected="${sheet == doc.activeSheet}" view="normal" zoomScale="140" zoomScaleNormal="140" zoomScalePageLayoutView="100"></sheetView></sheetViews>`;
    tableString += "<sheetData>";

    sheet.cells.sort(function(a, b) {
        let div = a.row - b.row;
        if (div == 0) div = a.col - b.col;
        return div;
    })
    let maxRow = sheet.cells.at(-1).row;
    for (let i = 1; i <= maxRow; i++) {
        tableString += `<row r="${i}">`;
        const rowCells = sheet.cells.filter(c => c.row == i);
        if (rowCells.length > 0) {

            let maxCol = rowCells.at(-1).col;
            for (let j = 1; j <= maxCol; j++) {
                const colCell = rowCells.find(c => c.col == j);
                if (colCell) {
                    let type = "";
                    let style = 0;
                    let value: any = "";
                    if (colCell.value) {
                        switch (colCell.type) {
                            case "string":
                                type = "s";
                                style = 1;
                                value = doc._strings.length;
                                doc._strings.push(colCell.value.toString())
                                break;
                            case "float":
                                type = "n";
                                style = 2;
                                value = colCell.value;
                                break;
                            case "date": // days since 30.12.1899 with decimal for time
                                type = "n";
                                style = 1;
                                let dateValue = colCell.value as Date
                                let timeZone = dateValue.getHours() - dateValue.getUTCHours()
                                let valueTime = dateValue.getTime() + (timeZone * T_HOUR)
                                value = (valueTime - T_1899_12_30) / T_DAY;
                                break;
                            default:
                                value = colCell.value;
                                break;
                        }
                    }

                    tableString += `<c r="${getCellName(colCell)}" s="${style}" t="${type}">`;
                    tableString += `<v>${escapeXML(value)}</v>`;
                    tableString += "</c>";
                }
            }
        }
        tableString += "</row>";
    }

    tableString += "</sheetData>";
    tableString += "</worksheet>";
    return tableString;
}