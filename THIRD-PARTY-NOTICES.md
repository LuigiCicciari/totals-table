# Third-party notices

Totals Table includes the following third-party software in the `lib` folder, unmodified. Each is used under its own license, and full license texts are in the `licenses` folder.

## Tableau Extensions API

- File: `lib/tableau.extensions.1.latest.min.js` (version 1.17)
- Source: https://github.com/tableau/extensions-api
- Copyright (c) Salesforce, Inc.
- License: Apache License 2.0, full text in [licenses/Apache-2.0-tableau-extensions-api.txt](licenses/Apache-2.0-tableau-extensions-api.txt)

## ExcelJS

- File: `lib/exceljs.min.js` (version 4.4.0, browser build)
- Source: https://github.com/exceljs/exceljs
- Copyright (c) 2014-2019 Guyon Roche
- License: MIT, full text in [licenses/MIT-exceljs.txt](licenses/MIT-exceljs.txt)

### Packages bundled inside the ExcelJS browser build

The ExcelJS browser build packages these open-source libraries into the same file. They were identified from the bundle's contents, and their licenses come from the npm registry. The versions inside the bundle may differ from the current releases, and each package's own repository has its full license text.

| Package | License |
|---|---|
| JSZip | MIT or GPL-3.0-or-later (used here under MIT) |
| pako | MIT and Zlib |
| saxes | ISC |
| xmlchars | MIT |
| dayjs | MIT |
| fast-csv | MIT |
| uuid | MIT |
| core-js | MIT |
| regenerator-runtime | MIT |
| readable-stream | MIT |
| buffer | MIT |
| base64-js | MIT |
| ieee754 | BSD-3-Clause |
| safe-buffer | MIT |
| events | MIT |
| inherits | ISC |
| string_decoder | MIT |
| util-deprecate | MIT |
| setimmediate | MIT |
| lie | MIT |
| immediate | MIT |
