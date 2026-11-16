# Changelog

All notable changes to Totals Table. Versions follow [Semantic Versioning](https://semver.org): MAJOR.MINOR.PATCH, where a patch fixes bugs, a minor version adds features without breaking anything, and a major version changes something existing (read its notes before updating). The version in use shows when hovering the info icon in the table's toolbar.

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
