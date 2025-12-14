# 🎨 Conditional Formatting - Complete Guide

> **Automatically format cells based on their values - make data visual and insights instant**

## 📋 Table of Contents

- [What is Conditional Formatting?](#what-is-conditional-formatting)
- [Getting Started](#getting-started)
- [Highlight Cell Rules](#highlight-cell-rules)
- [Top/Bottom Rules](#topbottom-rules)
- [Data Bars](#data-bars)
- [Color Scales](#color-scales)
- [Icon Sets](#icon-sets)
- [Custom Formula Rules](#custom-formula-rules)
- [Managing Rules](#managing-rules)
- [Real-World Examples](#real-world-examples)
- [Best Practices](#best-practices)

---

## What is Conditional Formatting?

### Definition
**Conditional Formatting** automatically changes cell formatting (color, font, borders, etc.) based on cell values or formulas.

### Why Use It?

**Benefits:**
- ✅ Visual data analysis at a glance
- ✅ Highlight important values automatically
- ✅ Spot trends and patterns
- ✅ Flag exceptions and outliers
- ✅ Create dynamic dashboards
- ✅ No formulas required (for basic rules)

**Common Uses:**
- Highlight overdue dates in red
- Show high/low values with colors
- Display progress bars in cells
- Traffic light indicators for KPIs
- Heat maps for data visualization
- Duplicate value detection

---

## Getting Started

### Accessing Conditional Formatting

**Location:** **Home** tab → **Conditional Formatting**

**Quick Access:** Select cells → **Home** → **Conditional Formatting** → Choose rule type

**Keyboard:** **Alt + H + L** (Windows)

---

## Highlight Cell Rules

### Overview
Highlight cells that meet specific criteria

### Greater Than
**Highlight values greater than a number**

**Steps:**
1. Select range
2. **Conditional Formatting** → **Highlight Cells Rules** → **Greater Than**
3. Enter value (or cell reference)
4. Choose format
5. Click **OK**

**Example:** Highlight sales over $1,000 in green
```
Range: B2:B100
Greater Than: 1000
Format: Light Green Fill
```

**Dynamic Reference:**
```
Greater Than: =$B$1
(Uses value in B1 as threshold - updates automatically)
```

---

### Less Than
**Highlight values less than a number**

**Use Cases:**
- Below average performance
- Low inventory alerts
- Under budget items

**Example:** Flag scores below 60 in red
```
Range: C2:C50
Less Than: 60
Format: Light Red Fill
```

---

### Between
**Highlight values in a range**

**Steps:**
1. Select range
2. **Between**
3. Enter: Minimum and Maximum
4. Choose format

**Example:** Highlight acceptable range
```
Between: 50 AND 100
Format: Yellow Fill
(Values outside range not highlighted)
```

---

### Equal To
**Highlight exact matches**

**Use Cases:**
- Find specific status
- Match target values
- Identify categories

**Example:** Highlight "Approved" status
```
Equal To: "Approved"
Format: Green Fill with Dark Green Text
```

---

### Text Contains
**Highlight cells containing specific text**

**Use Cases:**
- Find keywords
- Flag error messages
- Identify categories

**Example:** Highlight cells containing "Urgent"
```
Text that Contains: "Urgent"
Format: Red Fill with Dark Red Text
```

**Partial Match:**
Works with partial text:
```
Contains: "apple"
Matches: "Apple", "Pineapple", "apple pie"
```

---

### Date Occurring
**Highlight dates in specific timeframes**

**Options:**
- Today
- Tomorrow
- Yesterday
- Last 7 days
- Last week
- This week
- Next week
- Last month
- This month
- Next month

**Example:** Highlight upcoming deadlines
```
Date Occurring: Next 7 Days
Format: Orange Fill
```

**Use Cases:**
- Overdue items: Yesterday
- This week tasks: This Week
- Upcoming events: Next Week

---

### Duplicate Values
**Highlight duplicates or unique values**

**Steps:**
1. Select range
2. **Highlight Cells Rules** → **Duplicate Values**
3. Choose: **Duplicate** or **Unique**
4. Select format

**Example:** Find duplicate entries
```
Range: A2:A100
Highlight: Duplicate
Format: Light Red Fill
```

**Find Unique Values:**
```
Highlight: Unique
Format: Green Fill
(Only values that appear once)
```

---

## Top/Bottom Rules

### Top 10 Items
**Highlight highest values**

**Steps:**
1. Select range
2. **Top/Bottom Rules** → **Top 10 Items**
3. Enter number (default 10)
4. Choose format

**Example:** Top 5 salespeople
```
Top: 5 Items
Format: Green Fill with Dark Green Text
```

**Works with any number:**
```
Top 3, Top 20, Top 100, etc.
```

---

### Top 10 %
**Highlight top percentage**

**Example:** Top 20% of values
```
Top: 20 %
Format: Green Fill
```

**Use:** Performance rankings, percentile analysis

---

### Bottom 10 Items / Bottom 10 %
**Highlight lowest values or bottom percentage**

**Example:** Flag bottom performers
```
Bottom: 5 Items
Format: Red Fill
```

---

### Above Average / Below Average
**Highlight values above or below average**

**Automatic Calculation:**
Excel calculates average automatically

**Example:** Above average sales
```
Above Average
Format: Green Fill
```

**Dynamic:**
Average recalculates when data changes

---

## Data Bars

### Overview
**Visual bars inside cells** proportional to cell values

**Effect:**
```
Sales:
100  ████
50   ██
150  ██████
```

### Apply Data Bars

**Steps:**
1. Select range
2. **Conditional Formatting** → **Data Bars**
3. Choose color scheme

**Colors Available:**
- Gradient Fill (Blue, Green, Red, Orange, Light Blue, Purple)
- Solid Fill (Blue, Green, Red, Orange, Light Blue, Purple)

**Example:** Visualize sales performance
```
Range: B2:B20 (Sales column)
Data Bars: Gradient Fill - Blue
```

### Customize Data Bars

**Advanced Options:**
1. **Manage Rules** → Edit Rule
2. **Format Style**: Options
   - **Minimum/Maximum**: Automatic, Number, Percent, Formula, Percentile
   - **Bar Direction**: Context, Left-to-Right, Right-to-Left
   - **Negative Values**: Color and axis
   - **Bar Appearance**: Solid or Gradient, Border

**Show Bar Only (Hide Numbers):**
1. Edit Rule
2. **Show Bar Only**: Check

**Example:** Progress bars 0-100%
```
Minimum: Number 0
Maximum: Number 100
Show Bar Only: Yes
```

---

## Color Scales

### Overview
**Gradient color schemes** based on value ranges

**2-Color Scale:**
```
Low values:  Red
High values: Green
Middle:      Gradient
```

**3-Color Scale:**
```
Low:    Red
Mid:    Yellow
High:   Green
```

### Apply Color Scales

**Steps:**
1. Select range
2. **Conditional Formatting** → **Color Scales**
3. Choose scheme

**Presets:**
- Green-Yellow-Red
- Red-Yellow-Green
- Green-White-Red
- Red-White-Green
- Blue-White-Red
- And more...

**Example:** Heat map of sales
```
Range: B2:F20 (sales by month/product)
Color Scale: Red-Yellow-Green
```

### Custom Color Scales

**Create Custom:**
1. **Conditional Formatting** → **Color Scales** → **More Rules**
2. Set:
   - **Format Style**: 2-Color or 3-Color Scale
   - **Minimum**: Type, Value, Color
   - **Midpoint**: (3-color only)
   - **Maximum**: Type, Value, Color

**Example:** Custom scale
```
Minimum: Lowest Value - Red
Midpoint: Percentile 50 - Yellow
Maximum: Highest Value - Green
```

---

## Icon Sets

### Overview
**Icons** displayed based on value thresholds

**Available Icons:**
- Arrows (3, 4, 5 variations)
- Traffic lights
- Flags
- Symbols
- Ratings (stars)
- Indicators

### Apply Icon Sets

**Steps:**
1. Select range
2. **Conditional Formatting** → **Icon Sets**
3. Choose icon set

**Example:** Traffic lights for status
```
Range: E2:E50 (Performance scores)
Icon Set: 3 Traffic Lights
  Red: Bottom 33%
  Yellow: Middle 33%
  Green: Top 33%
```

### Customize Icon Sets

**Advanced Options:**
1. **Manage Rules** → Edit Rule
2. Customize:
   - Icon style
   - Reverse icon order
   - Show icon only
   - Custom thresholds

**Example:** Custom thresholds
```
Green: >= 90
Yellow: >= 70
Red: < 70
```

**Change Icon Order:**
```
Reverse: Green on bottom, Red on top
```

**Show Icons Only:**
```
Show Icon Only: Hide values, show icons
```

---

## Custom Formula Rules

### Create Rule with Formula

**Most Powerful Feature:**
Use ANY Excel formula to determine formatting

**Steps:**
1. Select range
2. **Conditional Formatting** → **New Rule**
3. Choose: **Use a formula to determine which cells to format**
4. Enter formula (must return TRUE/FALSE)
5. Click **Format**
6. Set format
7. Click **OK**

### Important Formula Rules

**Relative vs Absolute References:**
```
=$A1 = 100          // Formula adjusts per row
=$A$1 = 100         // Always checks A1
=A$1 = 100          // Row fixed, column relative
```

**Formula MUST return TRUE or FALSE**

---

### Common Formula Examples

**Alternate Row Colors**
```
Formula: =MOD(ROW(),2)=0
Format: Light Gray Fill
(Every even row colored)
```

**Highlight Entire Row Based on Column**
```
Range: A2:E100
Formula: =$B2="Complete"
Format: Green Fill
(Entire row green if column B = "Complete")
```

**Weekend Highlighting**
```
Range: A2:A100 (dates)
Formula: =OR(WEEKDAY(A2)=1, WEEKDAY(A2)=7)
Format: Light Blue Fill
```

**Highlight Overdue Tasks**
```
Range: A2:D100
Formula: =AND($C2<TODAY(), $D2<>"Complete")
Format: Red Fill
(Due date passed AND not complete)
```

**Highlight Duplicates in Column**
```
Range: A2:A100
Formula: =COUNTIF($A$2:$A$100, A2)>1
Format: Yellow Fill
```

**Compare Two Columns**
```
Range: C2:C100
Formula: =$B2<>$C2
Format: Orange Fill
(Highlight differences)
```

**Highlight Blanks**
```
Formula: =ISBLANK(A2)
Format: Red Fill
```

**Top 10% with Formula**
```
Formula: =A2>=PERCENTILE($A$2:$A$100, 0.9)
Format: Green Fill
```

**Alternate Column Colors**
```
Formula: =MOD(COLUMN(),2)=0
Format: Light Gray Fill
```

**Highlight Based on Text Length**
```
Formula: =LEN(A2)>50
Format: Yellow Fill
(Flag long entries)
```

**3-Color Risk Scale**
```
High Risk: =B2>100
  Format: Red
Medium: =AND(B2>=50, B2<=100)
  Format: Yellow
Low: =B2<50
  Format: Green
```

---

## Managing Rules

### View Rules
**See all conditional formatting rules**

**Access:** **Conditional Formatting** → **Manage Rules**

**Show Rules For:**
- This Worksheet
- Current Selection

**Rule Details:**
- Rule type
- Applies to range
- Format preview
- Stop if True checkbox

---

### Edit Rules

**Modify Existing Rule:**
1. **Manage Rules**
2. Select rule
3. Click **Edit Rule**
4. Make changes
5. Click **OK**

---

### Delete Rules

**Remove Rules:**
- **Clear Rules from Selected Cells**
- **Clear Rules from Entire Sheet**
- **Clear Rules from This Table**

**Or:**
1. **Manage Rules**
2. Select rule
3. Click **Delete Rule**

---

### Rule Precedence

**Multiple Rules on Same Cell:**
- Rules applied top to bottom
- Later rules can override earlier ones
- **Stop If True**: Stops checking remaining rules

**Reorder Rules:**
- Drag up/down in Manage Rules dialog

**Example:**
```
Rule 1: >100 = Green
Rule 2: >1000 = Red
Result: Values >1000 show Red (Rule 2 overrides)
```

---

### Copy Conditional Formatting

**Method 1: Format Painter**
1. Select cell with formatting
2. Click **Format Painter**
3. Select destination cells

**Method 2: Paste Special**
1. Copy cell (Ctrl+C)
2. Select destination
3. **Paste Special** → **Formats**

---

## Real-World Examples

### Project Status Dashboard
```
Status Column:
  "Complete" = Green
  "In Progress" = Yellow
  "Not Started" = Red
  "Blocked" = Orange

Use: Highlight Cell Rules → Equal To
```

### Sales Performance Heat Map
```
Monthly sales table
Color Scale: Red (low) → Yellow → Green (high)
Shows hot and cold months at glance
```

### Overdue Invoice Tracker
```
Formula: =AND(DueDate<TODAY(), Status<>"Paid")
Format: Red Fill with Bold Text
Highlights unpaid overdue invoices
```

### Grade Scoring
```
A: >=90 = Dark Green
B: >=80 = Light Green
C: >=70 = Yellow
D: >=60 = Orange
F: <60 = Red

Use: Icon Sets or multiple rules
```

### Budget Variance Analysis
```
Over Budget: Actual > Budget = Red
Under Budget: Actual < Budget = Green
Data Bars: Show variance magnitude
```

### Inventory Alerts
```
Low Stock: <10 = Red
Medium: 10-50 = Yellow
Good: >50 = Green

Use: Icon Sets (Traffic Lights)
```

---

## Best Practices

### Design Guidelines

**Color Usage:**
- ✅ Use color consistently
- ✅ Red = Bad/High Alert
- ✅ Yellow = Warning/Medium
- ✅ Green = Good/OK
- ✅ Limit colors (3-5 max)
- ✅ Consider colorblind users

**Clarity:**
- ✅ Don't over-format
- ✅ Key data only
- ✅ Make rules obvious
- ✅ Test on others

### Performance

**Large Datasets:**
- ✅ Limit conditional formatting range
- ✅ Use fewer rules
- ✅ Avoid complex formulas
- ✅ Turn off for printing if slow

### Maintenance

**Documentation:**
- ✅ Name rules descriptively
- ✅ Comment complex formulas
- ✅ Keep rules simple when possible
- ✅ Regular cleanup of unused rules

---

## Quick Reference

| Type | Use Case | Example |
|------|----------|---------|
| Greater Than | Values above threshold | Sales > $1000 |
| Text Contains | Find keywords | Contains "Urgent" |
| Duplicate Values | Find duplicates | Highlight duplicates |
| Data Bars | Visual comparison | Progress bars |
| Color Scales | Heat maps | Sales by region |
| Icon Sets | Status indicators | Traffic lights |
| Custom Formula | Complex logic | Overdue & incomplete |

---

**[⬆ Back to Main README](../../README.md)**
