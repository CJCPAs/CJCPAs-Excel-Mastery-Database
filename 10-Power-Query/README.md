# ⚡ Power Query - Complete Guide

> **Transform, clean, and combine data from any source with Excel's most powerful ETL tool**

## 📋 Table of Contents

- [What is Power Query?](#what-is-power-query)
- [Getting Started](#getting-started)
- [Data Sources](#data-sources)
- [The Power Query Editor](#the-power-query-editor)
- [Common Transformations](#common-transformations)
- [Combining Data](#combining-data)
- [Advanced Techniques](#advanced-techniques)
- [M Language Basics](#m-language-basics)
- [Best Practices](#best-practices)

---

## What is Power Query?

### Definition
**Power Query** is Excel's ETL (Extract, Transform, Load) tool that lets you **connect to, transform, and load data** from virtually any source without formulas.

### Why Use Power Query?

**Without Power Query:**
- Manual copy/paste
- Complex formulas
- Repetitive cleaning steps
- Difficult to update
- Error-prone

**With Power Query:**
- ✅ Automated data refresh
- ✅ No formulas needed
- ✅ Repeatable transformations
- ✅ Clean, transform, combine data
- ✅ Connect to any data source
- ✅ M language for advanced users

### Common Use Cases
- **Data Cleaning**: Remove duplicates, trim spaces, fix formatting
- **Data Transformation**: Pivot/unpivot, split columns, change types
- **Data Combination**: Merge tables, append datasets
- **Data Import**: Files, databases, web, APIs
- **Automation**: Refresh with one click
- **ETL Workflows**: Extract, transform, load pipelines

### Availability
- **Excel 2016+**: Built-in (Get & Transform)
- **Excel 2013/2010**: Download Power Query add-in
- **Excel 365**: Latest features always available

---

## Getting Started

### Opening Power Query

**Method 1: Import Data**
1. **Data** tab → **Get Data** (Excel 2016+)
2. Or **Get & Transform** group
3. Choose data source
4. Power Query Editor opens

**Method 2: From Table**
1. Select data or cell in table
2. **Data** tab → **From Table/Range**
3. Confirm table has headers
4. Click **OK**

**Method 3: Create Blank Query**
- **Data** → **Get Data** → **From Other Sources** → **Blank Query**

### The Basic Workflow

```
1. CONNECT → Choose data source
   ↓
2. TRANSFORM → Clean and shape data
   ↓
3. LOAD → Send to Excel worksheet or Data Model
   ↓
4. REFRESH → Update with one click
```

### Your First Power Query

**Example: Clean a messy table**

**Source Data:**
```
| Name         | Sales    |
|--------------|----------|
|  John Smith  | $1,234   |
| MARY JONES   | $ 987    |
|jane doe      |  $2,345  |
```

**Steps in Power Query:**
1. Import data: **Data** → **From Table/Range**
2. Editor opens with steps on right
3. Transformations:
   - Trim spaces: Right-click Name → **Transform** → **Trim**
   - Proper case: Right-click Name → **Transform** → **Capitalize Each Word**
   - Remove $: Right-click Sales → **Replace Values** → "$" with ""
   - Change type: Right-click Sales → **Change Type** → **Whole Number**
4. Click **Close & Load**

**Result:**
```
| Name        | Sales |
|-------------|-------|
| John Smith  | 1234  |
| Mary Jones  | 987   |
| Jane Doe    | 2345  |
```

**Future Updates:**
- Change source data
- Right-click query → **Refresh**
- All transformations re-apply automatically!

---

## Data Sources

### Files

**Excel Workbooks**
- **Data** → **Get Data** → **From File** → **From Workbook**
- Select file → Choose sheet or table
- Transform as needed

**CSV/Text Files**
- **Data** → **Get Data** → **From File** → **From Text/CSV**
- Auto-detects delimiters
- Change delimiter if needed

**Folder (Multiple Files)**
- **Data** → **Get Data** → **From File** → **From Folder**
- Combines all files in folder
- Perfect for: Monthly reports, daily logs

**Example: Combine all Excel files in folder**
```
1. Get Data → From Folder
2. Navigate to folder → OK
3. Click "Combine & Transform"
4. All files merge into one table!
```

### Databases

**SQL Server**
- **Data** → **Get Data** → **From Database** → **From SQL Server Database**
- Enter server name
- Choose database and table
- Import or write SQL query

**Access Database**
- **Data** → **Get Data** → **From Database** → **From Microsoft Access Database**
- Select .accdb or .mdb file
- Choose tables

**Other Databases**
- MySQL, PostgreSQL, Oracle, etc.
- May require ODBC drivers

### Web

**From Web**
- **Data** → **Get Data** → **From Other Sources** → **From Web**
- Enter URL
- Power Query detects tables on page
- Select table to import

**Example: Import stock prices, weather data, etc.**

**Web API**
- **Data** → **Get Data** → **From Other Sources** → **From Web**
- Enter API endpoint URL
- Parse JSON/XML

### Online Services

- **SharePoint**
- **OneDrive for Business**
- **Dynamics 365**
- **Salesforce**
- **Google Analytics**
- Many more connectors

---

## The Power Query Editor

### Interface Overview

**Left Pane: Queries**
- List of all queries
- Organize into groups
- Rename queries

**Center: Data Preview**
- Shows first 1000 rows (default)
- Column headers with type icons
- Profile, quality, distribution (optional)

**Right Pane: Query Settings**
- Query name
- **Applied Steps**: All transformations
- Properties

**Top Ribbon: Transformation Options**
- **Home**: Common transforms
- **Transform**: Column operations
- **Add Column**: Create new columns
- **View**: Display options

### Applied Steps

**Every action creates a step:**
```
Applied Steps:
  Source
  Promoted Headers
  Changed Type
  Trimmed Text
  Capitalized Each Word
  Replaced Value
```

**Edit Steps:**
- Click step to see that point in process
- Click X to delete step
- Click gear icon to edit settings
- Drag to reorder (careful!)

**Formula Bar:**
- Shows M code for selected step
- Edit directly for advanced users
- `fx` button to view

---

## Common Transformations

### Column Operations

**Rename Column**
- Double-click header
- Or right-click → **Rename**

**Remove Columns**
- Select column(s)
- Right-click → **Remove** or **Remove Other Columns**
- Or **Home** → **Remove Columns**

**Change Data Type**
- Click type icon in header
- Or right-click → **Change Type**
- Types: Text, Number, Date, True/False, etc.

**Move Columns**
- Drag column header
- Or right-click → **Move** → (Left/Right/To Beginning/To End)

---

### Text Transformations

**Clean Text**
```
Right-click column → Transform:
  • Trim: Remove extra spaces
  • Clean: Remove non-printable characters
  • Uppercase: ALL CAPS
  • Lowercase: all lowercase
  • Capitalize Each Word: Title Case
```

**Split Column**
```
Right-click column → Split Column:
  • By Delimiter: Comma, space, custom
  • By Number of Characters: Fixed width
  • By Positions: Specify positions
```

**Example: Split "Last, First" into two columns**
```
1. Right-click Name column
2. Split Column → By Delimiter
3. Choose: Comma
4. Split at: Each occurrence of delimiter
```

**Extract Text**
```
Right-click column → Extract:
  • Length: Number of characters
  • First Characters: From left
  • Last Characters: From right
  • Range: Specify start and length
  • Text Before/After Delimiter
```

**Replace Values**
```
Right-click column → Replace Values
  • Value to Find: "Old"
  • Replace With: "New"
```

---

### Number Transformations

**Mathematical Operations**
```
Right-click number column → Transform:
  • Multiply
  • Divide
  • Add
  • Subtract
  • Percentage
  • Absolute Value
  • Round/Round Up/Round Down
```

**Statistics**
```
Right-click number column → Statistics:
  • Sum
  • Average
  • Median
  • Min/Max
  • Count Values
```

**Standard**
```
• Scientific
• Percentage
• Integer
• Decimal Number
```

---

### Date Transformations

**Extract Date Parts**
```
Right-click date column → Date:
  • Year
  • Month (number or name)
  • Day
  • Quarter
  • Week of Year
  • Day of Week
  • Day of Year
```

**Date Calculations**
```
• Age (days from today)
• Add/Subtract Days
• Start/End of Month, Quarter, Year
```

**Example: Extract Month Name**
```
1. Right-click Date column
2. Date → Month → Name of Month
3. New column created: "December"
```

---

### Filter Rows

**Basic Filters**
```
Click dropdown in column header:
  • Uncheck items to hide
  • Search box to find
  • Select All / Clear All
```

**Text Filters**
```
• Equals / Does Not Equal
• Begins With / Ends With
• Contains / Does Not Contain
```

**Number Filters**
```
• Equals
• Greater Than / Less Than
• Between
• Greater Than or Equal To
```

**Date Filters**
```
• Before / After
• Between
• Today, Yesterday, This Week, etc.
```

**Remove Blank Rows**
```
Click dropdown → Remove Empty
```

**Top N / Bottom N**
```
Home → Keep Rows → Keep Top Rows
Enter number of rows
```

---

### Sort

**Sort Column**
- Click dropdown → **Sort Ascending/Descending**
- Or right-click → **Sort**

**Multi-Level Sort**
1. Sort by first column
2. Hold Shift
3. Sort by second column
4. Repeat for additional levels

---

### Group By (Aggregate)

**Summarize Data**
```
Home → Group By
```

**Settings:**
- **Group By**: Column(s) to group
- **New Column Name**: Name for result
- **Operation**: Sum, Average, Count, Min, Max, etc.
- **Column**: Column to aggregate

**Example: Total Sales by Product**
```
Group By:
  • Group by: Product
  • New column: Total Sales
  • Operation: Sum
  • Column: Sales
```

**Multiple Aggregations:**
Click **Add aggregation** for additional calculations

---

### Pivot & Unpivot

**Pivot Column**
```
Transform → Pivot Column
```
- **Values Column**: Column with values
- **Advanced**: Aggregate function

**Use:** Turn rows into columns

**Example:**
```
Before:
| Product | Month | Sales |
|---------|-------|-------|
| Apple   | Jan   | 100   |
| Apple   | Feb   | 150   |

After Pivot:
| Product | Jan | Feb |
|---------|-----|-----|
| Apple   | 100 | 150 |
```

**Unpivot Columns**
```
Select columns → Transform → Unpivot Columns
```

**Use:** Turn columns into rows (reverse of pivot)

**Example:**
```
Before:
| Product | Jan | Feb |
|---------|-----|-----|
| Apple   | 100 | 150 |

After Unpivot:
| Product | Attribute | Value |
|---------|-----------|-------|
| Apple   | Jan       | 100   |
| Apple   | Feb       | 150   |
```

**Perfect for:** Normalizing wide tables

---

## Combining Data

### Append Queries (Union)

**Combine rows from multiple tables**

**Requirements:**
- Same column structure
- Same column names (or similar)

**Steps:**
1. **Home** → **Append Queries**
2. Choose: **Two tables** or **Three or more tables**
3. Select tables to combine
4. Click **OK**

**Use Cases:**
- Combine monthly reports
- Merge files from multiple locations
- Stack similar datasets

**Example:**
```
Sales_Jan:
| Product | Sales |
|---------|-------|
| Apple   | 100   |

Sales_Feb:
| Product | Sales |
|---------|-------|
| Orange  | 200   |

After Append:
| Product | Sales |
|---------|-------|
| Apple   | 100   |
| Orange  | 200   |
```

---

### Merge Queries (Join)

**Combine columns from multiple tables**

**Types of Joins:**
- **Left Outer**: All from left, matching from right
- **Right Outer**: All from right, matching from left
- **Full Outer**: All from both
- **Inner**: Only matching rows
- **Left Anti**: Only in left, not in right
- **Right Anti**: Only in right, not in left

**Steps:**
1. Select first query
2. **Home** → **Merge Queries**
3. Select second query
4. Choose matching columns
5. Select join type
6. Click **OK**
7. Click expand icon in new column
8. Choose columns to add

**Example: Add Prices to Sales**
```
Sales:
| Product | Quantity |
|---------|----------|
| Apple   | 10       |

Prices:
| Product | Price |
|---------|-------|
| Apple   | $5    |

After Merge:
| Product | Quantity | Price |
|---------|----------|-------|
| Apple   | 10       | $5    |
```

---

## Advanced Techniques

### Custom Columns

**Add calculated column**
```
Add Column → Custom Column
```

**Examples:**
```m
// Full Name
= [FirstName] & " " & [LastName]

// Total
= [Quantity] * [Price]

// Conditional
= if [Sales] > 1000 then "High" else "Low"

// Date calculation
= Date.Year([OrderDate])
```

---

### Conditional Column

**Create column with if/then logic**
```
Add Column → Conditional Column
```

**Easier than Custom Column for simple conditions**

**Example:**
```
If [Sales] > 1000 then "High"
else if [Sales] > 500 then "Medium"
else "Low"
```

---

### Index Column

**Add row numbers**
```
Add Column → Index Column
```

**Options:**
- From 0
- From 1
- Custom (start, increment)

---

### Duplicate Column

**Create copy of column**
```
Right-click column → Duplicate Column
```

**Use:** Preserve original while transforming copy

---

### Replace Errors

**Handle errors in column**
```
Right-click column → Replace Errors
Enter replacement value
```

**Or:**
```m
= try [Column] otherwise null
```

---

## M Language Basics

### What is M?

**M** is the formula language of Power Query
- Similar to Excel formulas but more powerful
- Every transformation creates M code
- Can write custom M for advanced scenarios

### Viewing M Code

**See generated M:**
1. Click step in Applied Steps
2. Look at formula bar
3. Or **View** → **Advanced Editor**

### Basic M Syntax

**Let-In Expression:**
```m
let
    Source = Excel.Workbook(...),
    Sheet1 = Source{[Name="Sheet1"]}[Data],
    Result = Table.TransformColumnTypes(Sheet1, ...)
in
    Result
```

**Comments:**
```m
// Single line comment
/* Multi-line
   comment */
```

**Common M Functions:**

**Table Functions:**
```m
Table.SelectRows(table, condition)
Table.AddColumn(table, name, function)
Table.RemoveColumns(table, {"Column1"})
Table.RenameColumns(table, {{"Old", "New"}})
```

**Text Functions:**
```m
Text.Upper("hello")                → "HELLO"
Text.Trim(" hello ")               → "hello"
Text.Replace("hello", "e", "a")    → "hallo"
```

**Date Functions:**
```m
Date.Year(#date(2025,12,14))       → 2025
Date.AddMonths(#date(2025,1,1), 3) → 3/1/2025
```

**Conditional:**
```m
if [Sales] > 1000 then "High" else "Low"
```

---

## Best Practices

### Data Preparation

**Source Data:**
- ✅ Use Excel Tables (auto-expand)
- ✅ Consistent column headers
- ✅ No merged cells
- ✅ Data validation at source

### Query Organization

**Naming:**
- ✅ Descriptive query names
- ✅ Use folders/groups for organization
- ✅ Prefix with purpose (e.g., "Dim_Products", "Fact_Sales")

**Loading:**
- ✅ Load final tables to worksheet or model
- ✅ Intermediate queries: Right-click → **Enable Load** (uncheck)
- ✅ Keeps query list clean

### Performance

**Speed Up:**
- ✅ Filter early (reduce rows first)
- ✅ Remove columns you don't need
- ✅ Disable auto type detection if needed
- ✅ Use Table.Buffer for small lookup tables
- ✅ Fold queries (push to data source when possible)

**Query Folding:**
- Power Query pushes steps to data source (SQL, database)
- Faster than loading all data then transforming
- Check: Right-click step → **View Native Query**
- If available = folding works

### Error Handling

**Robust Queries:**
```m
= try [Column] otherwise null
= if [Column] = null then 0 else [Column]
```

**Data Type:**
- ✅ Set explicit data types
- ✅ Handle type errors

### Documentation

**Comment Steps:**
- Right-click step → **Properties**
- Add description
- Helps future you (and others)

---

## Common Use Cases

### Clean Messy Data
```
1. Import dirty data
2. Trim whitespace
3. Proper case
4. Remove duplicates
5. Fix data types
6. Replace errors
```

### Combine Monthly Files
```
1. Get Data → From Folder
2. Filter to desired files
3. Combine & Transform
4. Refresh monthly
```

### Unpivot Wide Table
```
1. Import data
2. Select ID columns
3. Transform → Unpivot Other Columns
4. Rename columns
```

### Build Date Table
```m
let
    StartDate = #date(2020,1,1),
    EndDate = #date(2030,12,31),
    DayCount = Duration.Days(EndDate - StartDate) + 1,
    Source = List.Dates(StartDate, DayCount, #duration(1,0,0,0)),
    TableFromList = Table.FromList(Source, Splitter.SplitByNothing()),
    ChangedType = Table.TransformColumnTypes(TableFromList,{{"Column1", type date}}),
    RenamedColumns = Table.RenameColumns(ChangedType,{{"Column1", "Date"}}),
    AddYear = Table.AddColumn(RenamedColumns, "Year", each Date.Year([Date])),
    AddMonth = Table.AddColumn(AddYear, "Month", each Date.Month([Date])),
    AddMonthName = Table.AddColumn(AddMonth, "Month Name", each Date.MonthName([Date]))
in
    AddMonthName
```

---

## Quick Tips

### Keyboard Shortcuts

| Action | Shortcut |
|--------|----------|
| Open Advanced Editor | Ctrl + Shift + M |
| Refresh Preview | Ctrl + R |
| Close & Load | Ctrl + Shift + Enter |
| Rename Query | F2 |

### Right-Click is Your Friend
- Most operations available via right-click
- Column header, cell, query name

### Use Example File
- Power Query works on sample
- Creates query structure
- Apply to full dataset later

---

**[⬆ Back to Main README](../../README.md)**
