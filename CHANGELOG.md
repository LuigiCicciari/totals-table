# Changelog

All notable changes to Totals Table. Versions follow [Semantic Versioning](https://semver.org): MAJOR.MINOR.PATCH, where a patch fixes bugs, a minor version adds features without breaking anything, and a major version changes something existing (read its notes before updating). The version in use shows when hovering the info icon in the table's toolbar.

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
