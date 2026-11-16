# Changelog

All notable changes to Totals Table. Versions follow [Semantic Versioning](https://semver.org): MAJOR.MINOR.PATCH, where a patch fixes bugs, a minor version adds features without breaking anything, and a major version changes something existing (read its notes before updating). The version in use shows when hovering the info icon in the table's toolbar.

## 1.3.5 (2026-11-16)

**Changed**
- The freeze line shows a hand cursor (closing while you drag it), so it reads as something to move rather than a column edge to stretch.

## 1.3.4 (2026-11-16)

**Fixed**
- The last frozen column can be resized again. Its resize grip sat under the freeze line's handle; the freeze line now keeps the edge and the resize grip sits just inside it.

## 1.3.3 (2026-11-16)

**Changed**
- Clearing a selection works the way it does in Tableau and spreadsheets: click the selected row again, click any empty space in the table (below the rows, right of the last column, or the bar above the table), or press Esc. Each clears the selected rows, their selection on the rest of the dashboard, and the highlighted cells. Copying also clears the cell highlight, since it has done its job.

## 1.3.2 (2026-11-16)

**Fixed**
- Viewers no longer see build details anywhere: thresholds that come from a parameter show its current value instead of its name; total descriptions use the column headers on screen; notes about blank totals and which rule applied are shown only while editing; the Compare list leaves out columns hidden by a parameter, and a difference column hides when either of its columns is hidden.
- Top or Bottom N: viewers who can't change the layout can still turn it off with its chip; Excel keeps the Other rows even when subtotal rows are left out of the file; running totals follow the rows as shown; Keep only and Exclude no longer offer a value from inside an Other row; Copy whole table keeps grouped order.
- Excel data bars now cover data rows only, so they also work with subtotals and Other rows.
- A sort on a hidden column no longer applies unseen.
- Group names containing quotation marks work in Add to group.
- In browsers that block storage, layouts saved during the session are no longer lost when Tableau refreshes the data.
- A second workbook that shares a sheet name and data source names keeps its viewer layouts when its fields change.

## 1.3.1 (2026-11-16)

**Changed**
- Turning on Top or Bottom N now always sorts the table by that column, so the rows read as a ranking even when the table was sorted by another column. A different sort can still be chosen afterwards.

## 1.3.0 (2026-11-16)

**Added**
- Hide when: extends the existing "(none)" auto-hide to any parameter and value. In Format → Columns & totals, a column can hide itself while a chosen parameter has one of the listed values, for example a True/False parameter "Show Discount" set to False, or a list parameter "View" set to "Summary". The column comes back as soon as the parameter changes, stays out of viewers' show/hide list and out of downloads while hidden, and its table calculation columns (such as % of total) hide with it. Matching ignores capitalization; several values can be separated by commas. Off until set.

## 1.2.0 (2026-11-16)

**Added**
- Top and bottom N: from a number column's menu, choose Show top or bottom (5, 10, 20, 50 or any number), for example the top 10 cities by Sales. The rest are combined into an "Other" row that totals them with the usual rules (Profit Ratio stays a true ratio) and opens with a click to show its rows. With subtotals, each group gets its own top rows and Other row. Grand totals and subtotals still include every row. Zeros and blanks never make the top. A chip above the table shows it is on and turns it off. Excel matches the screen with each Other row grouped and closed; CSV still has every row.

## 1.1.0 (2026-11-16)

**Added**
- Values shown last: in Format → Display, list catch-all values such as "Other, Unknown". In the default order they sort after everything else in their column, in any text column: a Region called "Unknown" comes last, and a city recorded as "Texas Other" comes last among the Texas cities. With subtotals, matching groups come last at their level. Values match as whole words, ignoring case ("Other" doesn't match "Mother"). Once someone sorts, they sort like every other row, unless "Keep them last even when the table is sorted" is ticked. Off until set.

## 1.0.0 (2026-11-16)

First release.

**Table**
- Fast scrolling for large tables; rows appear while the rest load
- Column drag, move, hide, resize and rename; frozen columns with a draggable freeze line
- Multi-column sorting (Shift-click or "Then sort" in a column's menu)
- Filter row, search, and right-click Keep only / Exclude, with chips to clear them and a Clear filters message when nothing matches
- Subtotals up to three levels deep, with Collapse all / Expand all per level
- Group headers over related columns, movable as a block
- Headers that follow parameters, and columns hidden when their parameter is set to "(none)"
- Links for web-address columns, or text that links to one
- Table calculation columns: % of total, rank, running total, and difference or % difference against another column
- Column stats for any column
- Right-click copy; Tableau tooltips and mark selection

**Totals**
- Totals and subtotals from rules imported into each workbook, so ratio columns total as a ratio of their summed parts; per-column overrides

**Formatting and style**
- Conditional formatting (thresholds, color scales, data bars) in colorblind-safe colors, with symbols so meaning never depends on color alone
- Table style: text size, row height, colors with contrast checks, lines, shading, totals and negative numbers
- Per-column alignment, bold and number format overrides

**Downloads**
- Excel with real numbers and formats, colors, hyperlinks, frozen panes, merged group headers and a collapsible outline for subtotals
- CSV with raw data rows only, protected against spreadsheet formulas

**Viewers**
- Layout changes saved in the viewer's own browser, plus named layouts with export and import
- Viewers never see parameter or field names; those are shown only while editing
