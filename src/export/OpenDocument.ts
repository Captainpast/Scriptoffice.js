import * as JSZip from "jszip";
import { OfficeDocument } from "../OfficeDocument";
import { getCellName, SpreadsheetDocument, SpreadsheetDocumentCell, SpreadsheetDocumentSheet, SpreadsheetDocumentStyle } from "../SpreadsheetDocument";

type CSpreadsheetDocument = SpreadsheetDocument & { _styles: CSpreadsheetDocumentStyle[] }
type CSpreadsheetDocumentCell = SpreadsheetDocumentCell & { style: CSpreadsheetDocumentStyle }
type CSpreadsheetDocumentStyle = SpreadsheetDocumentStyle & { _name: string, _type: string, _target: number, _stringified: string }

export interface SpreadsheetOptions {
    compressionLevel: number
}

export function escapeXML(value: any): string {
    if (value) {
        value = value.toString();
        value = value.replaceAll("&", "&amp;");
        value = value.replaceAll("'", "&apos;");
        value = value.replaceAll("\"", "&quot;");
        value = value.replaceAll(">", "&gt;");
        value = value.replaceAll("<", "&lt;");
    }
    return value;
}

export function dateToString(date: Date): string {
    let str = date?.toISOString()
    str = str?.split(".")[0]
    return str;
}

function create(mimetype: string, doc: OfficeDocument): JSZip {
    var zip = new JSZip();
    zip.file("mimetype", mimetype, { compression: "STORE" })
    zip.file("META-INF/manifest.xml", `<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">` +
        `<manifest:file-entry manifest:full-path="/" manifest:media-type="${escapeXML(mimetype)}"/>` +
        `<manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>` +
        `<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>` +
        `<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>` +
        `<manifest:file-entry manifest:full-path="settings.xml" manifest:media-type="text/xml"/>` +
        `</manifest:manifest>`)
    
        doc.date ??= new Date();
    zip.file("meta.xml", `<?xml version="1.0" encoding="UTF-8"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:xlink="http://www.w3.org/1999/xlink"><office:meta>` +
            `<meta:generator>ScriptOffice/1.0$Web_UnixLike ${escapeXML(doc.generator)}</meta:generator>` +
            `<meta:initial-creator>${escapeXML(doc.initialCreator)}</meta:initial-creator>` +
            `<meta:creation-date>${dateToString(doc.creationDate)}</meta:creation-date>` +
            `<dc:creator>${escapeXML(doc.creator)}</dc:creator>` +
            `<dc:date>${dateToString(doc.date)}</dc:date>` +
        `</office:meta></office:document-meta>`)

    return zip;
}

export async function spreadsheet(doc: CSpreadsheetDocument, options: SpreadsheetOptions): Promise<ArrayBuffer> {
    var zip = create("application/vnd.oasis.opendocument.spreadsheet", doc)

    spreadsheetMergeStyles(doc);

    // WIP style.xml
    zip.file("style.xml", `<?xml version="1.0" encoding="UTF-8"?><office:document-styles></office:document-styles>`)
    // content.xml 
    zip.file("content.xml", `<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:css3t="http://www.w3.org/TR/css3-text/" xmlns:grddl="http://www.w3.org/2003/g/data-view#" xmlns:xhtml="http://www.w3.org/1999/xhtml" xmlns:formx="urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:rpt="http://openoffice.org/2005/report" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:oooc="http://openoffice.org/2004/calc" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:ooow="http://openoffice.org/2004/writer" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" xmlns:tableooo="http://openoffice.org/2009/table" xmlns:drawooo="http://openoffice.org/2010/draw" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" xmlns:dom="http://www.w3.org/2001/xml-events" xmlns:field="urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:xforms="http://www.w3.org/2002/xforms">` +
        "<office:automatic-styles>" +
            doc._styles.map(s => spreadsheetCellStyle(s)).join("") +
        "</office:automatic-styles>" +
        "<office:body><office:spreadsheet>" + 
            doc.sheets.map(s => spreadsheetSheet(s)).join("") +
            "<table:database-ranges>" +
                doc.sheets.map(s => spreadsheetDatabase(s)).join("") +
            "</table:database-ranges>" +
        "</office:spreadsheet></office:body></office:document-content>")
    // settings.xml 
    zip.file("settings.xml", `<?xml version="1.0" encoding="UTF-8"?><office:document-settings xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:ooo="http://openoffice.org/2004/office" xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.3"><office:settings>` +
    `<config:config-item-set config:name="ooo:view-settings"><config:config-item-map-indexed config:name="Views"><config:config-item-map-entry><config:config-item-map-named config:name="Tables">` +
        doc.sheets.map(s => spreadsheetSheetSettings(s)).join("") +
    "</config:config-item-map-named></config:config-item-map-entry></config:config-item-map-indexed></config:config-item-set>"+
    "</office:settings></office:document-settings>")

    var file = await zip.generateAsync({
        type: "arraybuffer",
        platform: "UNIX",
        mimeType: "application/vnd.oasis.opendocument.spreadsheet",
        compression: "DEFLATE",
        compressionOptions: { level: options?.compressionLevel || 9 } })
    return file;
}

function spreadsheetMergeStyles(doc: CSpreadsheetDocument) {
    doc._styles = [];
    for (let i = 0; i < doc.sheets.length; i++) {
        for (let j = 0; j < doc.sheets[i].cells.length; j++) {
            const cell = doc.sheets[i].cells[j]
            if (cell._style) {
                let style = cell.style as CSpreadsheetDocumentStyle;
                style._stringified = JSON.stringify(style, spreadsheetMergeStylesJsonReplacer)

                let sameStyle = doc._styles.find(s => s._stringified == style._stringified);
                if (sameStyle) {
                    cell._style = sameStyle;
                } else {
                    style._name = "ce" + doc._styles.length;
                    style._type = "table-cell";
                    doc._styles.push(style)
                }

                // adding style for column and row
                if (style.columnWidth) {
                    let sameStyle = doc._styles.find(s => s._type == "table-column" && s._target == cell.col);
                    if (sameStyle) {
                        // adust columnWidth to the lager value
                        if (style.columnWidth > sameStyle.columnWidth) {
                            sameStyle.columnWidth = style.columnWidth;
                        }
                    } else {
                        let newStyle = new SpreadsheetDocumentStyle() as CSpreadsheetDocumentStyle;
                        newStyle._name = "co" + doc._styles.length;
                        newStyle._type = "table-column";
                        newStyle._target = cell.col;
                        newStyle.columnWidth = style.columnWidth;
                        doc._styles.push(newStyle)
                    }
                    delete style.columnWidth;
                }
                if (style.rowHeight) {
                    let sameStyle = doc._styles.find(s => s._type == "table-row" && s._target == cell.row);
                    if (sameStyle) {
                        // adust rowHeight to the lager value
                        if (style.rowHeight > sameStyle.rowHeight) {
                            sameStyle.rowHeight = style.rowHeight;
                        }
                    } else {
                        let newStyle = new SpreadsheetDocumentStyle() as CSpreadsheetDocumentStyle;
                        newStyle._name = "ro" + doc._styles.length;
                        newStyle._type = "table-row";
                        newStyle._target = cell.row;
                        newStyle.rowHeight = style.rowHeight;
                        doc._styles.push(newStyle)
                    }
                    delete style.rowHeight;
                }
            }
        }
    }
}
function spreadsheetMergeStylesJsonReplacer(key: string, value: any) {
    const exclude = [ "columnWidth", "rowHeight" ]
    if (exclude.includes(key)) {
        return undefined;
    }
    return value;
}

function spreadsheetCellStyle(style: CSpreadsheetDocumentStyle): string {
    var styleString = `<style:style style:name="${escapeXML(style._name)}" style:family="${escapeXML(style._type)}">`

    if (style.bold || style.italic || style.color || (style.underline && style.underline != "none")) {
        styleString += "<style:text-properties";
        if (style.bold) styleString += ` fo:font-weight="bold"`;
        if (style.italic) styleString += ` fo:font-style="italic"`;
        if (style.color) styleString += ` fo:color="${style.color}"`;
        if (style.underline) styleString += ` style:text-underline-style="${escapeXML(style.underline)}" style:text-underline-width="auto" style:text-underline-color="${escapeXML(style.underlineColor)}"`;
        styleString += "/>";
    }

    if (style.backgroundColor) styleString += `<style:table-cell-properties fo:background-color="${escapeXML(style.backgroundColor)}"/>`;

    if (style.columnWidth) styleString += `<style:table-column-properties style:column-width="${escapeXML(style.columnWidth.toString())}cm"/>`;
    if (style.rowHeight) styleString += `<style:table-row-properties style:row-height="${escapeXML(style.rowHeight.toString())}cm"/>`;

    styleString += "</style:style>";
    return styleString;
}

function spreadsheetSheet(sheet: SpreadsheetDocumentSheet): string {
    var tableString = `<table:table table:name="${escapeXML(sheet.title)}">`
    
    if (sheet.cells.length > 0) {
        let doc = sheet.parentDocument as CSpreadsheetDocument;

        let colStyles = doc._styles.filter(s => s._type == "table-column").sort((a, b) => a._target - b._target);
        let maxCol = colStyles.at(-1)?._target;

        for (let i = 1; i <= maxCol; i++) {
            const colStyle = colStyles.find(s => s._target == i);
            if (colStyle) {
                tableString += `<table:table-column table:style-name="${colStyle._name}"/>`
            } else {
                let nextCol = colStyles.find(s => s._target > i)._target;
                tableString += `<table:table-column table:number-columns-repeated="${nextCol - i}"/>`;
                i = nextCol-1;
            }
        }

        sheet.cells.sort(function(a, b) {
            let div = a.row - b.row;
            if (div == 0) div = a.col - b.col;
            return div;
        })
        let maxRow = sheet.cells.at(-1).row;
        for (let i = 1; i <= maxRow; i++) {
            const rowCells = sheet.cells.filter(c => c.row == i);
            if (rowCells.length > 0) {
                tableString += "<table:table-row";
                let rowStyle = doc._styles.find(s => s._type == "table-row" && s._target == i);
                if (rowStyle) tableString += ` table:style-name="${rowStyle._name}"`;
                tableString += ">";

                let maxCol = rowCells.at(-1).col;
                for (let j = 1; j <= maxCol; j++) {
                    const colCell = rowCells.find(c => c.col == j);
                    if (colCell) {
                        tableString += spreadsheetCell(colCell as any);
                    } else {
                        const nextCol = rowCells.find(c => c.col > j).col;
                        tableString += `<table:table-cell table:number-columns-repeated="${nextCol - j}"/>`;
                        j = nextCol-1;
                    }
                }
                tableString += "</table:table-row>";
            } else {
                const nextRow = sheet.cells.find(c => c.row > i).row;
                tableString += `<table:table-row table:number-rows-repeated="${nextRow - i}"/>`;
                i = nextRow-1;
            }
        }
    }

    tableString += "</table:table>";
    return tableString;
}

function spreadsheetDatabase(sheet: SpreadsheetDocumentSheet): string {
    var tableString = "";

    for (let i = 0; i < sheet._databases.length; i++) {
        const database = sheet._databases[i];
        
        tableString += `<table:database-range table:name="${"_"+ i +"_"+ escapeXML(sheet.title)}" table:display-filter-buttons="true"`;

        tableString += ` table:target-range-address="&apos;${escapeXML(sheet.title)}&apos;.${getCellName(database.range.from)}`;
        tableString += `:&apos;${escapeXML(sheet.title)}&apos;.${getCellName(database.range.to)}">`;
    
        tableString += "</table:database-range>";
    }

    return tableString;
}

function spreadsheetSheetSettings(sheet: SpreadsheetDocumentSheet): string {
    var tableString = `<config:config-item-map-entry config:name="${escapeXML(sheet.title)}">`

    if (sheet._freezePos != null) {
        tableString += `<config:config-item config:name="HorizontalSplitMode" config:type="short">2</config:config-item>`;
        tableString += `<config:config-item config:name="VerticalSplitMode" config:type="short">2</config:config-item>`;

        tableString += `<config:config-item config:name="HorizontalSplitPosition" config:type="int">${sheet._freezePos.col - 1}</config:config-item>`;
        tableString += `<config:config-item config:name="VerticalSplitPosition" config:type="int">${sheet._freezePos.row - 1}</config:config-item>`;
        tableString += `<config:config-item config:name="PositionRight" config:type="int">${sheet._freezePos.col - 1}</config:config-item>`;
        tableString += `<config:config-item config:name="PositionBottom" config:type="int">${sheet._freezePos.row - 1}</config:config-item>`;
    }

    tableString += "</config:config-item-map-entry>";
    return tableString;
}

function spreadsheetCell(cell: CSpreadsheetDocumentCell): string {
    var cellString = `<table:table-cell office:value-type="${cell.type}" table:style-name="${cell.style._name}" `;
    if (cell.value instanceof Date) {
        cellString += `office:date-value="${dateToString(cell.value)}">`
    } else {
        cellString += `office:value="${escapeXML(cell.value) || ""}">`
    }
    cellString += `<text:p>${escapeXML(cell.value?.toString()) || ""}</text:p>`
    cellString += "</table:table-cell>"
    return cellString;
}