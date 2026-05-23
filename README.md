# 📊 Excel Skills Practical Assignment — Superstore Sales Dataset

A hands-on Excel project covering 12 core skill areas, built on the **Sample Superstore Sales** dataset from Tableau Community. Each section is a dedicated worksheet containing live formulas, structured data, and step-by-step task guides.

-----

## 📁 Project Structure

```
📁 Excel-Practical-Assignment/
├── 📄 README.md
├── 📊 Excel_Assignment.xlsx          # Main workbook — 12 section sheets
└── 📑 Excel_Module_Assignment.pdf    # Original assignment brief
```

**Sheets inside the workbook:**

|Sheet                       |Topic                                                               |
|----------------------------|--------------------------------------------------------------------|
|Section1 - Basics           |Arithmetic functions, TRANSPOSE, COUNT variants, YoY & CAGR         |
|Section2 - Formatting       |Cell formats, number formats, Wrap Text, Format Painter             |
|Section3 - Sort & Filter    |Number/date/text/color sorting, multi-level sort                    |
|Section4 - Dates            |DATEDIF, DATE, TODAY, YEAR/MONTH/DAY, ship speed classifier         |
|Section5 - Logical          |IF, AND, OR, Nested IF, IFERROR on real sales rep data              |
|Section6 - Data Cleaning    |UPPER/LOWER/PROPER/TRIM, SEARCH vs FIND, CONCATENATE, Go To Special |
|Section7 - Pivot Table      |Source data + step-by-step Pivot Table task guide                   |
|Section8 - Lookups          |VLOOKUP, HLOOKUP, INDEX/MATCH with a named Product DB               |
|Section9 - Cond. Aggregation|COUNTIF/SUMIF/AVERAGEIF and their multi-criteria IFS variants       |
|Section10 - Cond. Formatting|Color rules, color scales, data bars, icon sets, formula-based rules|
|Section11 - Consolidation   |Multi-year SUM consolidation, CONSOLIDATE feature, SUBTOTAL         |
|Section12 - VBA Macros      |VBA concepts, Developer tab guide, recorded & written macro examples|

-----

## 🔢 What Was Built — Section by Section

### Section 1: Basics of Excel

Monthly 2017 sales data with a full summary statistics panel: **SUM, MAX, MIN, SMALL (2nd lowest), LARGE (3rd highest), MEDIAN, AVERAGE, COUNT**. Includes a TRANSPOSE row using array formula, a COUNT/COUNTA/COUNTBLANK demo with intentionally blank cells, and a YoY growth table for 2014–2017 with **CAGR calculated using the POWER formula** (result: 14.8% compound growth over 3 years).

### Section 2: Formatting

Top-15 high-value orders formatted with currency, conditional highlighting, borders, and a Margin % status column. A dedicated number format reference table shows the format codes for percentages, thousands (K), negative-in-red, scientific notation, and accounting format. Wrap Text and Merge & Center applied to product description rows.

### Section 3: Sort & Filter

20 real orders pre-sorted descending by Sales with color tags (🟢 HIGH). Includes a tip panel covering single-level sort, 2-level sort (Category A→Z then Sales largest→smallest), sort by cell color, and alphabetical sort — demonstrated with the top-10 A→Z product names from the full dataset.

### Section 4: Dates

15 real order/ship date pairs showing duration calculated by both subtraction and DATEDIF (in days, months, years). A Ship Speed classifier using IF labels orders as ⚡ Express (≤2 days), ✈ Standard (3–4 days), or 🐢 Slow (5+ days). A reference table documents 11 date functions with syntax and example output.

### Section 5: Logical Functions

20 sales rep records with columns for: IF (High/Low vs $5,000 threshold), AND (sales ≥ target AND region = East), OR (sales ≥ 5,000 OR region = West), Nested IF (High/Mid/Low tiers), and IFERROR wrapping a Sales÷Target ratio. A quick-reference syntax table covers IF, AND, OR, NOT, Nested IF, and IFERROR.

### Section 6: Data Cleaning

12 customer names with deliberate extra spaces demonstrating UPPER, PROPER, LOWER, TRIM, LEN (raw vs trimmed), LEFT(5), RIGHT(4), and MID(4,3) — all in one table. A SEARCH vs FIND comparison shows case sensitivity differences across 5 text strings. CONCATENATE and & operator comparison table, plus a Go To Special and Find & Replace tips panel.

### Section 7: Pivot Table

50-row source dataset (Order ID through Profit) with a 6-task guide embedded in the sheet: basic pivot by Category, grouping by Month, multiple cross-tab tables (Region × Category), blank cell handling, conditional formatting on values, and a bonus custom group (Furniture + Office Supplies → “Non-Tech”).

### Section 8: Lookups

A 15-product **named range “ProductDB”** with 8 columns (ID, Name, Category, Sub-Category, Price, Stock, Reorder Level, Lead Days). VLOOKUP demonstrations cover exact match (FALSE), approximate match (TRUE), and IFERROR handling — including two intentional INVALID lookup IDs to show error trapping. HLOOKUP applied to a transposed 5-column version of the same data. Bonus INDEX/MATCH section showing left-side lookup that VLOOKUP cannot do. Comparison summary table: VLOOKUP vs HLOOKUP vs INDEX/MATCH.

### Section 9: Conditional Aggregation

30-order dataset used for 9 COUNTIF/SUMIF/AVERAGEIF formulas (single criterion) and 7 COUNTIFS/SUMIFS/AVERAGEIFS formulas (multiple criteria). Examples include: orders in East + Technology (0 — a real data insight), profitable Consumer orders (21 of 30), West + Technology total sales ($2,122.62), and South Furniture total profit (−$121.54, loss-making). Full syntax reference table at the bottom.

### Section 10: Conditional Formatting

30-order dataset with a 5-task guide: (1) highlight Sales > $500 in green, (2) 3-color scale on Sales (green→yellow→red), (3) gradient data bars on Profit (negative bars go left in red), (4) 3-arrow icon set on Quantity (≥5 green, ≥2 yellow, else red), (5) formula-based row rule flagging high-sales loss-makers using `=AND($E6<0,$D6>500)`. Manual color legend for the Performance Tag column (⭐ High / 〽️ Mid / 🔻 Low).

### Section 11: Consolidation

4-year monthly sales table (2014–2017) with Grand Totals — used to demonstrate cross-sheet SUM consolidation. Category × Region cross-tab with row % and analyst notes (e.g., Technology has best avg sale value; Furniture drag from loss-making Bookcases). SUBTOTAL demo with function codes 1/4/5/9 (AVG/MAX/MIN/SUM) versus a plain SUM — with annotation that SUBTOTAL respects active filters while SUM does not.

### Section 12: VBA Macros

Structured explanation of VBA concepts, advantages/disadvantages table, step-by-step Developer tab activation guide, and a full macro walkthrough (Record → Run → View Code → Edit → Save as .xlsm). Seven ready-to-paste VBA code snippets: Hello World, Format Header (RGB colors), Auto-Fit Columns, Alternate Row Colors (the loop used in this workbook itself), Input Box Greeting, Workbook_Open auto-run, and Write Numbers with squares. Security risks panel covering macro settings, digital signatures, and code red flags like `Shell` and `WScript.Run`.

-----

## 🛠️ Tools & Technologies

- **Microsoft Excel** (2016 / Microsoft 365 — 365 recommended for dynamic arrays)
- **VBA / Visual Basic for Applications** — Section 12
- **Sample Superstore Sales dataset** — [Tableau Community source](https://community.tableau.com/s/question/0D54T00000CWeX8SAL/sample-superstore-sales-excelxls)

> ⚠️ Save as `.xlsm` to preserve macros. Excel Online does not support VBA.

-----

## ⚡ Challenges Faced

**TRANSPOSE as an array formula** — `=TRANSPOSE()` must be entered with Ctrl+Shift+Enter in older Excel versions. The entire output range must be selected first, otherwise you get only the first value. In Excel 365 it spills automatically, which behaves differently and can confuse learners switching between versions.

**VLOOKUP with approximate match (TRUE)** — The source table must be sorted ascending by the lookup column, otherwise TRUE mode returns wrong results silently with no error. This caused incorrect price lookups until the ProductDB was sorted by Product ID.

**DATEDIF’s undocumented nature** — `DATEDIF` doesn’t appear in Excel’s formula autocomplete and has no tooltip. The `"YM"` and `"MD"` interval arguments especially require external documentation since Excel’s built-in help omits this function entirely.

**Blank cells breaking Pivot Table grouping** — A single blank in the Order Date column prevents Excel from grouping dates by Month/Quarter. Each blank had to be identified and filled before grouping worked correctly.

**SUBTOTAL vs SUM with filters** — The visual result looks identical when no filter is applied, which makes it easy to miss the difference. The distinction only becomes obvious when you apply a filter — at that point SUM ignores the filter while SUBTOTAL responds to it. This needed a side-by-side live demonstration to make the behavior clear.

**Macro security pop-ups** — Running the first macro triggered a security warning that blocked execution. The fix (File → Options → Trust Center → Macro Settings → “Disable all macros with notification”) is not obvious to first-time users and isn’t explained in most beginner tutorials.

-----

## 💡 Insights Gained

The **SUBTOTAL function** is significantly more practical than SUM for any filtered dataset — it’s the difference between a static total and a dynamic one that responds to whatever the user is currently viewing.

**INDEX/MATCH** solves the fundamental limitation of VLOOKUP (lookup column must be leftmost) and is resilient to column insertions that break VLOOKUP’s column index number. For any production workbook, INDEX/MATCH is the safer default.

**IFERROR wrapping** should be the default practice around VLOOKUP and any formula that can return `#N/A` or `#DIV/0!`. Without it, a single missing lookup ID cascades into broken dependent formulas across the whole sheet.

Building the **SEARCH vs FIND comparison** made the case-sensitivity difference immediately visible — SEARCH found “sales” inside “Superstore Sales Report” while FIND returned an error because it expected exact case. This is easy to forget in practice and can cause hard-to-debug formula behavior.

The **formula-based conditional formatting rule** (`=AND($E6<0,$D6>500)`) is far more powerful than the built-in highlight rules — it can evaluate conditions across multiple columns and apply formatting to an entire row, not just the cell being tested. The `$` column lock with relative row reference is the key syntax pattern to get right.

**CAGR via POWER** gives a more honest picture of growth than year-over-year percentages — the 2014–2017 Superstore data showed a 14.8% CAGR even though one year (2015) had negative growth, which point-to-point YoY comparisons would have overstated as a problem.

-----

## 🔧 Recommendations for Improvement

**Add INDEX/MATCH as a primary topic, not a bonus** — it appears here as a bonus in Section 8 but should be taught alongside VLOOKUP from the start. In professional Excel work, INDEX/MATCH or XLOOKUP (Excel 365) is preferred over VLOOKUP for robustness.

**Include XLOOKUP** for anyone on Excel 365 — it combines the best of VLOOKUP and INDEX/MATCH into a single, cleaner function that works in both directions and returns arrays.

**Power Query deserves its own section** — for data cleaning tasks in Section 6, Power Query is dramatically faster and more repeatable than manual formula-based cleaning. It handles large datasets that would slow down formula-heavy sheets.

**Charts and dashboards are a gap** — the dataset used throughout this project is rich enough to build meaningful visuals (category sales bar chart, monthly trend line, profit scatter plot), but chart creation isn’t covered. A Section 13 on charts would make the project feel complete.

**Use Excel Tables (Ctrl+T) from the start** — none of the data ranges here are converted to structured Tables. Tables auto-expand formulas, support structured references like `Table1[Sales]` instead of `D:D`, and make Pivot Tables update automatically when new rows are added.

**AVERAGEIFS is missing from the assignment brief** — it’s used in Section 9 (correctly) but wasn’t listed as a required task. Worth formalizing since it completes the IFS family alongside COUNTIFS and SUMIFS.

-----

## 📬 Acknowledgements

Assignment brief designed by **Shiva Vashishtha**  
Mentor LinkedIn: [linkedin.com/in/shivavashishtha](https://www.linkedin.com/in/shivavashishtha/)  
Dataset: Sample Superstore Sales — Tableau Community