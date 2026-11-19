# Totals Table (Tableau viz extension)

A fast table for Tableau worksheets, with correct totals for ratio columns, headers that follow parameters, and clean Excel or CSV downloads. It is plain HTML and JavaScript with no build step. Data never leaves the viewer's browser.

## Features

**Table**
- Columns: drag to reorder, move first/left/right/last, hide, show/hide picker, resize (double-click a header's right edge to fit the contents), rename/reset name
- Rows load A to Z by the text columns on the left, like a crosstab. Click a header to sort (text A to Z first, numbers largest first). Shift-click more headers, or use Then sort in a column's menu, to sort by several columns; headers show the sort order as small numbers. Viewers can always sort
- Optional values shown last (Format → Display): catch-all values like "Other" or "Unknown" sort after everything else in their column, in any of the text columns on the left that set the default order (such as Region, State, City). A Region called "Unknown" comes last, and a city recorded as "Texas Other" comes last among the Texas cities. Matching is by whole word, ignoring case, so "Other" doesn't match "Mother". Sorting then treats them like any other row, unless the author ticks "Keep them last even when the table is sorted"
- Optional filter row under the headers (text: contains, starts with or exact; numbers: min/max) and optional search box, both turned on in Format → Display. Search highlights the matching text and shows how many rows match
- Frozen header and totals row. Parameter-driven columns are frozen by default; drag the freeze line or use "Freeze up to here"
- Optional subtotals by any text column, up to three levels deep (for example Region, then State, then City): use Subtotal by this column, then Add subtotal level on the next column. Every level uses the totals rules, so ratio subtotals stay correct. Click a group row to collapse it. Collapse all in a level's column menu shows that level's subtotals, and Expand all opens it again. Sorting a number column orders the groups at every level by their subtotal and the rows within each group; sorting a subtotalled column orders that level by name
- Collapsed groups are remembered: what a viewer leaves collapsed stays that way on their next visit (without counting as a layout change), saved layouts keep their collapsed groups when saved or updated, and a whole level that was collapsed applies to new groups too. Authors can tick "Viewers start with this level collapsed" in a subtotal column's menu to set the starting view for everyone
- Links: columns holding web addresses are clickable automatically and open in a new tab. This includes parameter-driven columns: when a parameter switches a column to a web-address field, it becomes clickable, and switches back to plain text afterwards. In Format → Columns & totals, a link column can show a fixed label (such as "View") instead of the address, or any text column can link to a web-address column, so a name opens its page while the address column stays hidden. Only http and https addresses become links
- Group headers: a shared label over related columns, such as "Sales" over Sales, Profit and Profit Ratio, and "Orders" over Quantity and Discount. Set them in Format → Columns & totals, or with New group / Add to group in a column's menu. Drag a group header to move the whole group; it drops only between groups, never inside one, and the table scrolls when you drag near its edge. The group's menu also has Move group left, right, to start and to end. Membership is fixed by the author, so a column moved out of its group keeps its label above it in the new spot. Viewers can turn group headers off from the group's menu or the show/hide columns picker. Excel gets merged group header cells
- % of total: any column that sums can get a companion "% of total" column from its menu. Shares are of the filtered total, and subtotal rows show each group's share
- % of group: with subtotals on, a column that sums can also get a "% of group" column. Each row shows its share of its own group (for example each Category's share of its Region), and each subtotal row its share of the level above, so each group adds up to 100%. Without subtotals it shows the share of the total
- Top or bottom N from a number column's menu (5, 10, 20, 50 or any number): for example the top 10 cities by Sales, with the rest combined into an "Other" row that totals them with the usual rules and opens to show them. With subtotals, each group gets its own top rows and Other row. Totals still include every row, zeros and blanks never make the top, and a chip turns it off. Excel matches the screen; CSV keeps every row
- Table calculation columns from a number column's menu, each placed next to its column and downloaded like any other:
  - Rank (largest = 1 or smallest = 1) across the filtered rows; ties share a rank. Zeros always rank last, whichever direction is best, since a zero usually means no activity (such as $0 Sales for a city with no orders); blanks get no rank
  - Running total, in the order rows are shown (for columns that sum)
  - Compare with another column: difference (this − that) or % difference ((this − that) ÷ that). Totals and subtotals compare the two columns' own totals, so ratio columns compare correctly
- Keep only / Exclude: right-click a cell (or a range in one column) to filter the table to those values or remove them. A chip above the table shows each one; click it to remove. Like the filter row, these only affect the table, never Tableau's filters, and clear when the page reloads
- Column stats from any column's menu: rows, blanks, sum (only for columns that total by summing), average, median, min and max (with the row they come from) and the table total for numbers; distinct values and the most common for text; earliest, latest and distinct days for dates. Stats cover the filtered rows only, never totals or subtotals, and can be copied
- Right-click to copy cells, cells with headers, or the whole table (data rows only, in on-screen order). Tableau Desktop keeps Ctrl+C for itself, so right-click is the reliable way
- Tableau tooltips on hover, mark selection on click (Ctrl/Cmd-click adds), so dashboard actions keep working. Click the row again, click any empty space in the table, or press Esc to clear the selection
- Only the rows on screen are drawn, so large tables scroll smoothly. Rows appear while the rest are still loading
- The info icon shows the version Tableau is running, such as 1.0.0 (so you can tell when an update has reached Cloud; see [CHANGELOG.md](CHANGELOG.md) for what changed in each version), and load timing, split into Tableau query time and load/draw time
- Hovering a header shows its name, how it totals when that isn't a plain sum (for example "Profit ÷ Sales" for Profit Ratio, taken from the totals rules), and its sort. The column menu adds the total, any formatting rule and the original field name
- When filters leave no rows, the table says so and offers Clear filters

**Totals**
- Ratio columns total as a ratio of the summed parts, not a sum of ratios. Divide by zero gives a blank
- Weighted averages (opt-in, Format → Display → "Let viewers total a column as a weighted average"): a column menu's "Total as weighted average…" totals that column as Σ(value × weight) ÷ Σ(weight) using a chosen number column, for example Discount weighted by Sales. A rule of thumb: weight by whatever the number is "per". The choice is saved with each person's layout (an author's choice is the starting point for viewers), and the tooltip names the weight. Weights that add up to zero or less leave the total blank
- Totals and subtotals follow the table's own filters
- Totals come from rules imported into each workbook (see below), then per-column overrides: Auto, Sum, Average, Min, Max, Ratio (A ÷ B × multiplier) or None

**Conditional formatting** (Format → Formatting)
- Thresholds: good / watch / bad bands for any number column, with "higher is better" or "lower is better". Cut-off values can be fixed or come from a parameter, so a goal set on the dashboard recolors the table live
- Color scales: a gradient across the 5th to 95th percentile of what's on screen, good-to-bad or a single color
- Data bars: a bar behind each number, sized to the largest value on screen
- Accessible by design: the colors are colorblind-safe (checked under protan, deutan and tritan simulation), text contrast stays above 5:1, and thresholds also show ✓ ! ✕ symbols plus screen-reader text, so meaning never depends on color alone
- Totals and subtotals stay plain unless turned on per column. Formatting recalculates with filters
- Viewers who find it distracting can choose Hide formatting from a column's menu. Like other viewer changes, it's saved in their own browser and undone by Reset, and their downloads match what they see
- Excel keeps the colors and symbols (the cells stay real numbers), and data bars become Excel's own data bars

**Style** (Format → Style)
- The whole table's look: text size, row height (compact, normal, roomy), text color, header background and text color, bold or wrapped headers, row and column lines, alternating row shading, totals and subtotal colors and bold, how negative numbers show (-1,234, (1,234), or either in red), and an accent color for selection and links
- Everything starts at the original design, so nothing changes until it's set. Colors come with a short palette of readable choices or any #RRGGBB code, and each one shows its contrast, warning below 4.5:1
- Per-column overrides: alignment, bold, number style (number, currency, percent), decimals, units (K, M, B, or Auto from the column's largest value) and negatives. Difference and running total columns follow their source column's format
- Excel downloads follow the header, totals and subtotal colors, conditional formatting colors, bold, alignment and number formats (red negatives use Excel's own red); text size and row height are screen only

**Headers from parameters**
- If a field and a parameter share a name (for example a calculated field "Column 1" driven by a parameter "Column 1"), the header shows the parameter's current value and updates live. Parameters copied into sets with a dotted prefix also match: "Column 1" finds "east.Column 1" when there is no plain "Column 1". A prefix counts only when at least two parameters share it, so a lone "Target.Sales" never takes over a "Sales" column. Once a prefixed parameter is chosen for any column under Header from parameter, Auto only uses that prefix. If several prefixed sets fit and none is chosen, Auto links nothing rather than guess. The extension can't see calculation formulas, only names, so a parameter with the same name as the field always wins, even if the formula uses a different one. Format shows which parameter each column follows
- A column whose parameter is set to "(none)" is hidden automatically
- Hide when: any column can also hide itself while a chosen parameter has one of a list of values, for example a True/False parameter "Show Discount" set to False, or a list parameter "View" set to "Summary". Set it in Format → Columns & totals. It updates live, keeps the column out of viewers' show/hide list and downloads while hidden, and the column's % of total, rank or other calculation columns hide with it
- Any column can be linked to a different parameter, or unlinked, in Format → Columns & totals

**Downloads**
- Excel, CSV or both (both shows a small menu)
- What is on screen: visible columns, in order, with the current sort and filters
- Links come out as clickable hyperlinks in Excel and as plain text or addresses in CSV
- Excel: plain values, never formulas; Tableau's number formats carried over ($, %, thousands, decimals, K/M/B); real dates, including text columns that hold ISO dates (such as `STR([Date])`); bold header and totals; frozen header and frozen columns
- CSV: raw data only. Every row, never totals or subtotals. UTF-8, numbers at full value (not rounded to what the table shows) so pivots and recalculated ratios match the dashboard, the same values the Excel download stores, dates as `yyyy-mm-dd`
- Collapsed subtotal groups never hide rows from a download. In Excel, subtotal rows sit above their rows, grouped with Excel's +/- outline (one outline level per subtotal level, so Excel's 1 2 3 buttons collapse to each level), and groups collapsed in the table start collapsed in the file
- File name template with `{sheet}`, `{date}`, `{time}` and `{param:Parameter Name}`

**Viewer layouts**
- When "Let viewers move, hide, resize and freeze columns" is on, a viewer's changes (column order, hidden columns, widths, names, freeze, sort, subtotals, % of total columns, hidden formatting, hidden group headers and weighted-average choices) are saved in their own browser only. A "Custom layout · Reset" chip appears so they can return to the published layout
- Saved layouts: viewers can save named layouts (for example "Weekly review" or "Top regions") from the Layouts button and switch between them, update or undo changes to one, rename or delete it, or go back to the published layout. The chip shows which layout is active and whether it has been edited. A saved layout also remembers the parameters that decide what columns show (those linked to a column header or used by Hide when rules) and sets them again when it is chosen; date ranges and other parameters are never saved or changed
- Shared layouts: authors can save named layouts in the workbook from the Layouts button while editing (for example "Summary" and "Detail"). Every viewer sees them in their Layouts menu under Shared, above their own, can switch to them, and can save their own copies. Shared layouts carry the same column arrangement, column parameters and collapsed groups, and authors can export them or import a layouts file into them. Viewers can't change or delete shared layouts. Shared layouts need "Let viewers move, hide, resize and freeze columns" turned on
- Layouts live in the viewer's browser only, and the Layouts menu says so plainly: private windows, other computers and clearing cookies and site data won't have them. Export layouts saves them all to a small file that Import brings back on any browser or computer, or shares with a teammate. The first save shows a one-time tip about exporting, and if a browser blocks saving, the menu explains that too

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
- If a name fits more than one number column (for example `SUM(Sales)` and `AVG(Sales)` both show as Sales), the one SUM among them is used; otherwise the total is left blank until you choose its fields in Format. Naming the full field (`AVG(Sales)`) always picks that one
- The first matching rule wins, so put exceptions above general rules
- Both parts of a ratio must be on the Detail shelf. They can be hidden in the table and still feed the total. If one is missing, the total is blank and Format shows why
- `type` can be `sum`, `avg`, `min`, `max`, `ratio` or `none`

## Adding it to a worksheet

Tableau adds a viz extension from a `.trex` manifest that points to where the extension is hosted. Download [TotalsTable.trex](TotalsTable.trex) from this repository (it points to the hosted copy on GitHub Pages, so there is nothing to install). In a worksheet, open the Marks card type menu, choose **Add Extension** → **Access Local Viz Extensions**, and pick the `.trex`. Tableau copies the extension's address into the workbook, so the `.trex` is only needed when adding it.

Drag fields onto **Detail** in the order you want the columns. Fields that are only on **Tooltip** are left out of the table. Use **Format Extension** on the Marks card, or the gear icon, to open the settings. Settings are saved in the workbook.

To swap an existing table viz extension for this one, duplicate the worksheet and change the Marks card type on the copy. Detail-shelf fields carry over; column widths, order and names need setting again.

On Tableau Cloud, a site admin must add the extension's address to the site's allowed extensions.

## Status

Tested in Tableau Desktop and Tableau Cloud: loading and field carry-over, parameter headers and hide-when rules, freezing, single and multi-column sorting, filters, search and Keep only / Exclude, column dragging, totals and up to three levels of subtotals, collapsing, top and bottom N with Other rows, values shown last, % of total, rank, running total and compare columns, column stats, links, conditional formatting, table style and column formats, group headers, right-click copy, Excel and CSV downloads, viewer layouts and saved layouts with export and import, Tableau tooltips, and mark selection.

## Versions

Versions follow Semantic Versioning (MAJOR.MINOR.PATCH). [CHANGELOG.md](CHANGELOG.md) lists what changed in each one.

## Libraries

- Tableau Extensions API 1.17 (Apache 2.0)
- ExcelJS 4.4.0 (MIT), loaded only when someone downloads Excel

See [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md) for copyright notices and license texts, including the packages bundled inside ExcelJS.
