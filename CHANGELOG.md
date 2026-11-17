# Changelog

All notable changes to Totals Table. Versions follow [Semantic Versioning](https://semver.org): MAJOR.MINOR.PATCH, where a patch fixes bugs, a minor version adds features without breaking anything, and a major version changes something existing (read its notes before updating). The version in use shows when hovering the info icon in the table's toolbar.

## 1.6.1 (2026-11-17)

**Fixed**
- Double-clicking a column edge now shrinks a column that is wider than it needs to be, and repeated double-clicks no longer widen it. Link columns fit their link text instead of the web address behind it.
- Search matches what cells show: a link labeled "View" is no longer found by its hidden address, and matches inside link text are highlighted.
- A weighted average whose weight column is hidden goes back to the normal total, instead of naming the hidden column.
- Copying cells: text that starts like a spreadsheet formula (=, +, - or @) gets a leading apostrophe, so pasting into Excel can't run it. Number columns are copied as before.
- Table settings saved in the workbook are checked when the table loads: colors must be real hex colors, and layout values must be well-formed.
- A blank-total explanation in a header tooltip uses a colon instead of a dash.

## 1.6.0 (2026-11-17)

**Added**
- Double-click the right edge of a column header to fit the column to its contents (header, totals, subtotals and the rows shown, up to 600px wide).
- Search highlights the matching text in each cell, and the row count says how many rows match.
- % of group: a column that sums can get a companion "% of group" column from its menu while the table has subtotals. Each row shows its share of the group it sits in, and each subtotal row shows its share of the level above. Without subtotals it shows the share of the total, and its tooltip says so.
- Weighted averages, opt-in: with "Let viewers total a column as a weighted average" turned on in Format → Display, anyone who can change the layout can pick "Total as weighted average…" in a number column's menu and choose the weight column, for example Discount weighted by Sales. Totals and subtotals use Σ(value × weight) ÷ Σ(weight), the column menu and tooltip say which weight is used, and the choice is saved with that person's layout. Rows missing either value are left out, and weights that add up to zero or less leave the total blank. "Use the normal total" goes back.

**Fixed**
- The column resize grip is easier to grab: all of it is now inside the header, and the menu arrow no longer sits under it.

## 1.5.0 (2026-11-16)

**Added**
- Shared layouts: while editing, authors can save named layouts in the workbook from the Layouts button. Every viewer finds them in the Layouts menu under Shared, above their own layouts, and can switch to them or save their own copy. Authors can update, rename, delete, export and import them; viewers can't change them.
- Collapsed groups are remembered. What a viewer leaves collapsed stays that way on their next visit, without counting as a layout change. Saved and shared layouts keep their collapsed groups, and a level collapsed as a whole also collapses groups that appear in new data. "Update" is always offered for the layout in use, so its current view can be saved at any time.
- Authors can tick "Viewers start with this level collapsed" in a subtotal column's menu to set the starting view for everyone. Collapsing while editing no longer changes what viewers see by accident.

## 1.4.2 (2026-11-16)

**Fixed**
- Saved layouts only ever set the parameters that decide what columns show today, even if a layout file lists others, so a layout can never change a date range.
- Parameters added or renamed after the table was placed are picked up without reopening the workbook (in Format and for Hide when and layouts).
- Dragging the freeze line while the table is scrolled sideways moves it where you point instead of jumping several columns. The freeze line and the resize grip beside it are easier to grab.
- "Reset name" for viewers only appears for their own renames and brings back the author's name.
- Layout actions (Update, Rename, Delete, choosing a layout) work correctly even if the data refreshes while the Layouts menu is open. Choosing a layout no longer briefly shows "(edited)", and sorting then clearing the sort no longer counts as an edit.
- "Save current layout as" with an existing name asks before replacing it.
- Selections and highlighted cells clear when the data refreshes or the table is re-sorted, instead of landing on different rows.
- Data loads run one at a time, so quick successive refreshes can't collide.
- Excel data bars respect a viewer's Hide formatting and scale from zero like the bars on screen.
- Copying a subtotal or Other row copies what it shows. CSV header names are protected against spreadsheet formulas.
- Tableau tooltips close when the pointer moves onto subtotal or Other rows. A resize can no longer trigger a sort by accident. Right-click allows paste in text boxes.
- "Keep them last even when sorted" checks the same leading text columns as the default order.

## 1.4.1 (2026-11-16)

**Fixed**
- A saved layout that doesn't have its column settings yet (such as one saved before 1.4.0) now shows as edited, so "Update" appears in the Layouts menu and saves them in one click.

## 1.4.0 (2026-11-16)

**Added**
- Saved layouts remember the parameters that decide what columns show: parameters linked to a column header (for example a "Column 1" parameter set to Category or Region) and parameters used by Hide when rules. Switching to a layout sets them again, along with its column arrangement, so a "Category View" and a "Region View" can each bring back their own columns. Date ranges and other parameters are never saved or changed. Changing one of those parameters by hand marks the layout as edited, and Update saves the new choice. A saved value that is no longer available is skipped with a short note. Export and import include these values. Layouts saved before this version pick them up the next time they're updated.

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
