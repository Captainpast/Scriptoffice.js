_scriptOffice for JavaScript_

<div align="center">
  <img src="https://gitlab.com/Captainpast/scriptoffice.js/-/raw/main/logo.svg" alt="scriptOffice.js">
</div>

a library to create files of an office software suite like [LibreOffice](https://www.libreoffice.org/) programmcity and automated

# Using

## NPM / YARN / PNPM
run one of the command in your nodejs project
```bash
$ npm i script-office
$ yarn add script-office
$ pnpm add script-office
```
add to file
```js
import { OfficeDocument } from "script-office";
```

## CDN
add to html head
```html
<script src="https://cdn.jsdelivr.net/npm/script-office/dist/script-office.min.js"></script>
```

# Documentation

## Spreadsheet

### creating a new document

```js
var doc = OfficeDocument.create({ type: "spreadsheet" });
// or
var doc = OfficeDocument.create("spreadsheet");

// get the active and only sheet of the document
var sheet = doc.activeSheet;
```

### importing an existing document
currently not supported

### working with sheets
a spreadsheet can have multiple sheets
``` js
// get the current and active sheet
var sheet = doc.activeSheet;

// create sheet
var sheet = doc.addSheet("first");

// get sheet by index
var sheet = doc.getSheet(1);
// get sheet by name 
var sheet = doc.getSheet("first");

// set title
sheet.title = "first";
```

### working with cells
single cell
```js
// get cell
var cell = sheet.getCell("A1")
var cell = sheet.getCell({ col: "A", row: 1 })
var cell = sheet.getCell({ col: 1, row: 1 })

// set value
cell.value = "test";
cell.value = 2000;
cell.value = new Date();

// change style
cell.style.bold = true;
cell.style.italic = true;
cell.style.underline = "solid";
cell.style.underlineColor = "#ff0000";
cell.style.color = "reb(255,0,0)";
cell.style.backgroundColor = "red";
cell.style.columnWidth = 20;
cell.style.rowHeight = 10;

// at once
sheet.setCell("A1", {
  value: "test",
  style: {
    bold: true,
    underline: "dash"
  }
})
```
multiple cells
```js
// get cells content
var content = sheet.getCells("A1:B3").map(c => c.value)

// set cells
sheet.setCells("A1:B3", { style: { backgroundColor: "#ff0000" } })
```

### specials
```js
sheet.freezeAt("A1")
sheet.autoFilter("A1", "K9")
```

### export
there a several export formats:

#### OpenDocument spreadsheet
```js
var data = await doc.export("ods", {
    compressionLevel: 9
  });
fs.writeFileSync(path.join(__dirname, "test.ods"), Buffer.from(data))
```

#### Comma-separated values
only one sheet of the SpreadsheetDocument can be exported
```js
var data = await doc.export("csv", {
    separator: ";",
    useQuotes: true,
    sheet: "name or index of the sheet"
});
fs.writeFileSync(path.join(__dirname, "test.csv"), data)
```


## Build with webpack

cloning git repo and install dependencies
```bash
$ git clone https://gitlab.com/Captainpast/scriptoffice.js
$ cd scriptoffice.js
$ npm install
```

run `dist` script to create a minified and compact version
```bash
$ npm run dist
```
The release is now in the `./dist` folder and can be imported.

```js
import { OfficeDocument } from "./dist/script-office.min.js";
```
or
```html
<script src="./dist/script-office.min.js"></script>
```

# Roadmap
- xlsx export
- html export
- imports
- texts, drawings, presentations