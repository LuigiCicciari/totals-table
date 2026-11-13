# Totals Table (Tableau viz extension)

A fast table for Tableau worksheets, with correct totals for ratio columns, headers that follow parameters, and clean Excel or CSV downloads. It is plain HTML and JavaScript with no build step. Data never leaves the viewer's browser.

## Features

**Table**
- Columns: drag to reorder, move first/left/right/last, hide, show/hide picker, resize, rename/reset name
- Rows load A to Z by the text columns on the left, like a crosstab. Click a header to sort (text A to Z first, numbers largest first); viewers can always sort
- Per-column filters (text: contains, starts with or exact; numbers: min/max), optional search box
- Frozen header and totals row. Parameter-driven columns are frozen by default; drag the freeze line or use "Freeze up to here"
- Optional subtotals by any text column. Click a group row to collapse it, or use Collapse all / Expand all in that column's menu. Sorting a number column orders the groups by their subtotal and the rows within each group; sorting the subtotalled column orders the groups by name
- Right-click to copy cells, cells with headers, or the whole table (data rows only, in on-screen order). Tableau Desktop keeps Ctrl+C for itself, so right-click is the reliable way
- Tableau tooltips on hover, mark selection on click (Ctrl/Cmd-click adds), so dashboard actions keep working
- Only the rows on screen are drawn, so large tables scroll smoothly. Rows appear while the rest are still loading
- The info icon shows load timing, split into Tableau query time and load/draw time

**Totals**
- Ratio columns total as a ratio of the summed parts, not a sum of ratios. Divide by zero gives a blank
- Totals and subtotals follow the table's own filters
- Totals come from rules imported into each workbook (see below), then per-column overrides: Auto, Sum, Average, Min, Max, Ratio (A ÷ B × multiplier) or None

**Headers from parameters**
- If a field and a parameter share a name (for example a calculated field "Column 1" driven by a parameter "Column 1"), the header shows the parameter's current value and updates live
- A column whose parameter is set to "(none)" is hidden automatically
- Any column can be linked to a different parameter, or unlinked, in Format → Columns & totals

**Downloads**
- Excel, CSV or both (both shows a small menu)
- What is on screen: visible columns, in order, with the current sort and filters
- Excel: plain values, never formulas; Tableau's number formats carried over ($, %, thousands, decimals, K/M/B); real dates, including text columns that hold ISO dates (such as `STR([Date])`); bold header and totals; frozen header and frozen columns
- CSV: raw data only. Every row, never totals or subtotals. UTF-8, numbers rounded to what the table shows, dates as `yyyy-mm-dd`
- Collapsed subtotal groups never hide rows from a download. In Excel, subtotal rows sit above their rows, grouped with Excel's +/- outline, and groups collapsed in the table start collapsed in the file
- File name template with `{sheet}`, `{date}`, `{time}` and `{param:Parameter Name}`

**Viewer layouts**
- When "Let viewers move, hide, resize and freeze columns" is on, a viewer's changes are saved in their own browser only. A "Custom layout · Reset" chip appears so they can return to the published layout

## Totals rules

Rules live in each workbook's extension settings. Keep your master rules file somewhere private and import it with Format → Totals rules → Import rules file. Nothing about your rules is stored in this repository.

```json
{
  "version": 1,
  "rules": [
    { "when": { "equals": "Profit Ratio" },  "total": { "type": "ratio", "numerator": "Profit", "denominator": "Sales" } },
    { "when": { "endsWith": " per Order" },  "total": { "type": "ratio", "numerator": "{rest}", "denominator": "Orders" } },
    { "when": { "equals": "Discount" },      "total": { "type": "avg" } }
  ],
  "default": { "type": "sum" }
}
```

This example uses fields from Tableau's Superstore sample data, where `Orders` is a calculated field for `COUNTD([Order ID])`. With these rules, Profit Ratio totals as total Profit ÷ total Sales, "Sales per Order" and "Profit per Order" total as their measure ÷ Orders, Discount averages, and everything else sums.

- `when` takes one of `equals`, `startsWith`, `endsWith` or `contains`, matched against the field name without its aggregation (`AGG(Profit Ratio)` is matched as `Profit Ratio`), ignoring case
- `{rest}` is the rest of the name after the matched part: `Sales per Order` with `endsWith: " per Order"` gives `Sales`
- Column lookups accept singular or plural, so `Order` finds an `Orders` column
- The first matching rule wins, so put exceptions above general rules
- Both parts of a ratio must be on the Detail shelf. They can be hidden in the table and still feed the total. If one is missing, the total is blank and Format shows why
- `type` can be `sum`, `avg`, `min`, `max`, `ratio` or `none`

## Adding it to a worksheet

Tableau adds a viz extension from a `.trex` manifest that points to where the extension is hosted. In a worksheet, open the Marks card type menu, choose **Add Extension** → **Access Local Viz Extensions**, and pick the `.trex`. Tableau copies the extension's address into the workbook, so the `.trex` is only needed when adding it.

Drag fields onto **Detail** in the order you want the columns. Fields that are only on **Tooltip** are left out of the table. Use **Format Extension** on the Marks card, or the gear icon, to open the settings. Settings are saved in the workbook.

To swap an existing table viz extension for this one, duplicate the worksheet and change the Marks card type on the copy. Detail-shelf fields carry over; column widths, order and names need setting again.

On Tableau Cloud, a site admin must add the extension's address to the site's allowed extensions.

## Status

Tested in Tableau Desktop and Tableau Cloud: loading and field carry-over, parameter headers, freezing, sorting, filters, column dragging, totals and subtotals, collapsing, right-click copy, Excel and CSV downloads, and viewer layouts. Still to confirm: whether the Tableau tooltip clears on mouse-out.

## Libraries

- Tableau Extensions API 1.17 (MIT)
- ExcelJS 4.4.0 (MIT), loaded only when someone downloads Excel
