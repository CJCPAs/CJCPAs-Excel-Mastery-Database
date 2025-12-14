# ✅ Data Validation - Complete Guide

> **Control data entry, prevent errors, and ensure data quality with Excel's powerful validation tools**

## 📋 Table of Contents

- [What is Data Validation?](#what-is-data-validation)
- [Getting Started](#getting-started)
- [Validation Types](#validation-types)
- [Input Messages](#input-messages)
- [Error Alerts](#error-alerts)
- [Advanced Techniques](#advanced-techniques)
- [Real-World Examples](#real-world-examples)
- [Best Practices](#best-practices)

---

## What is Data Validation?

### Definition
**Data Validation** controls what data can be entered into a cell by creating rules that must be followed.

### Why Use Data Validation?

**Benefits:**
- ✅ Prevent data entry errors
- ✅ Ensure consistency
- ✅ Create dropdown lists
- ✅ Limit value ranges
- ✅ Validate dates
- ✅ Custom rules with formulas
- ✅ Improve data quality

**Common Uses:**
- Dropdown selection lists
- Date range restrictions
- Number limits (age, quantity, price)
- Text length limits
- Custom formula validation
- Prevent duplicates
- Required fields

---

## Getting Started

### Accessing Data Validation

**Location:** **Data** tab → **Data Validation**

**Quick Access:** 
1. Select cell(s)
2. **Data** → **Data Validation**
3. Configure settings

**Keyboard:** **Alt + D + L** (Windows)

---

## Validation Types

### List (Dropdown)

**Create Dropdown List**

**Method 1: Type Items**
```
Allow: List
Source: Apple,Orange,Banana
```

**Method 2: Reference Range**
```
Allow: List
Source: =$A$1:$A$10
```

**Method 3: Named Range**
```
Allow: List
Source: =Products
```

**Method 4: Dynamic Table Column**
```
Allow: List
Source: =Table1[ProductName]
```

**Example: Status Dropdown**
```
Settings:
  Allow: List
  Source: Not Started,In Progress,Complete,On Hold
  In-cell dropdown: ✓
```

**Tips:**
- ✅ Separate items with commas (no spaces unless needed)
- ✅ Use named ranges for flexibility
- ✅ Table columns auto-expand
- ✅ Reference can be on different sheet

**Dynamic Dropdown (Based on Another Cell):**
```
Cell A2: Category dropdown (Fruit, Vegetable)
Cell B2: =INDIRECT(A2)
Named Ranges: Fruit (Apple, Orange), Vegetable (Carrot, Broccoli)
```

---

### Whole Number

**Restrict to Integer Values**

**Operators:**
- Between
- Not between
- Equal to
- Not equal to
- Greater than
- Less than
- Greater than or equal to
- Less than or equal to

**Example 1: Age Validation**
```
Allow: Whole number
Data: between
Minimum: 18
Maximum: 100
```

**Example 2: Quantity > 0**
```
Allow: Whole number
Data: greater than
Minimum: 0
```

**Example 3: Even Numbers Only**
```
Allow: Custom
Formula: =MOD(A1,2)=0
```

---

### Decimal

**Allow Decimal Numbers**

**Example: Price Validation**
```
Allow: Decimal
Data: between
Minimum: 0
Maximum: 1000
```

**Example: Percentage (0-1)**
```
Allow: Decimal
Data: between
Minimum: 0
Maximum: 1
```

---

### Date

**Restrict Date Ranges**

**Example 1: Future Dates Only**
```
Allow: Date
Data: greater than
Start date: =TODAY()
```

**Example 2: Date Range**
```
Allow: Date
Data: between
Start date: 1/1/2024
End date: 12/31/2024
```

**Example 3: Workdays Only**
```
Allow: Custom
Formula: =AND(A1>=TODAY(), WEEKDAY(A1,2)<=5)
(Future dates, Monday-Friday only)
```

**Example 4: Within 30 Days**
```
Allow: Date
Data: between
Start date: =TODAY()
End date: =TODAY()+30
```

---

### Time

**Validate Time Values**

**Example: Business Hours**
```
Allow: Time
Data: between
Start time: 9:00 AM
End time: 5:00 PM
```

**Example: After Specific Time**
```
Allow: Time
Data: greater than or equal to
Start time: 8:00 AM
```

---

### Text Length

**Limit Characters**

**Example 1: Maximum Length**
```
Allow: Text length
Data: less than or equal to
Maximum: 50
```

**Example 2: Exact Length (ZIP Code)**
```
Allow: Text length
Data: equal to
Length: 5
```

**Example 3: Range**
```
Allow: Text length
Data: between
Minimum: 8
Maximum: 20
(Password requirements)
```

**Example 4: SSN Format with Custom**
```
Allow: Custom
Formula: =AND(LEN(A1)=11, ISNUMBER(--SUBSTITUTE(A1,"-","")))
(Format: 123-45-6789)
```

---

### Custom (Formula)

**Most Powerful - Use Any Formula**

**Formula Must Return TRUE or FALSE**

**Example 1: No Duplicates**
```
Allow: Custom
Formula: =COUNTIF($A$1:$A$100, A1)=1
```

**Example 2: Email Validation**
```
Allow: Custom
Formula: =AND(LEN(A1)>0, ISNUMBER(FIND("@",A1)), ISNUMBER(FIND(".",A1)))
```

**Example 3: Uppercase Only**
```
Allow: Custom
Formula: =EXACT(A1, UPPER(A1))
```

**Example 4: Must Start With Specific Text**
```
Allow: Custom
Formula: =LEFT(A1,3)="ABC"
```

**Example 5: Depends on Another Cell**
```
Allow: Custom
Formula: =A1>B1
(Value in A must be greater than B)
```

**Example 6: Required Field (No Blanks)**
```
Allow: Custom
Formula: =LEN(A1)>0
```

**Example 7: Numeric Text (Phone Numbers)**
```
Allow: Custom
Formula: =ISNUMBER(--A1)
(Text that looks like number)
```

**Example 8: Date Must Be Weekend**
```
Allow: Custom
Formula: =OR(WEEKDAY(A1)=1, WEEKDAY(A1)=7)
```

---

## Input Messages

### Purpose
**Show helpful message when cell is selected**

### Configure Input Message

**Settings Tab:**
```
☑ Show input message when cell is selected
Title: Enter Product Name
Input message: Choose a product from the dropdown list
```

**Appearance:**
- Yellow tooltip-style message
- Appears when cell selected
- Guides user on what to enter

**Example:**
```
Title: Required Field
Message: Please enter your full name (First Last)
```

**Tips:**
- ✅ Keep short and clear
- ✅ Explain what's expected
- ✅ Mention format if specific
- ✅ Be helpful, not condescending

---

## Error Alerts

### Purpose
**Show message when invalid data entered**

### Alert Styles

**Stop (Default)**
```
Icon: ⊗ Red X
Effect: Prevents entry
Options: Retry or Cancel
Use: Strict validation
```

**Warning**
```
Icon: ⚠ Yellow Triangle
Effect: Warns but allows entry
Options: Yes, No, Cancel
Use: Soft enforcement
```

**Information**
```
Icon: ℹ Blue i
Effect: Informs but allows entry
Options: OK, Cancel
Use: Guidance only
```

### Configure Error Alert

**Error Alert Tab:**
```
☑ Show error alert after invalid data is entered
Style: Stop
Title: Invalid Entry
Error message: Please enter a number between 1 and 100.
```

**Examples:**

**Stop Alert (Strict):**
```
Style: Stop
Title: Invalid Product
Message: Product not found in list. Please select from dropdown.
```

**Warning Alert (Soft):**
```
Style: Warning
Title: Unusual Value
Message: This value seems high. Are you sure it's correct?
```

**Information Alert:**
```
Style: Information
Title: Tip
Message: Recommended range is 10-50. You can enter other values if needed.
```

---

## Advanced Techniques

### Dependent Dropdowns

**Dropdown changes based on another cell**

**Setup:**
```
Categories in A1: Fruit, Vegetable, Dairy
Named Ranges:
  Fruit = {Apple, Orange, Banana}
  Vegetable = {Carrot, Broccoli, Spinach}
  Dairy = {Milk, Cheese, Yogurt}

Cell A2: List = Categories
Cell B2: =INDIRECT(A2)
```

**How it works:**
1. Select category in A2
2. B2 dropdown shows only items for that category

**Example: State → City**
```
State list in A2
City list in B2: =INDIRECT(A2)
Named ranges: Texas={Houston,Dallas}, California={LA,SF}
```

---

### Multi-Level Dependent Dropdowns

**3+ Levels Deep**

**Example: Country → State → City**
```
Country: USA, Canada
State (if USA): Texas, California
State (if Canada): Ontario, Quebec
City: Depends on State
```

**Implementation:**
- Use INDIRECT with concatenation
- Or nested INDIRECT formulas
- Named ranges for each combination

---

### Dynamic Lists

**List Grows/Shrinks Automatically**

**Method 1: Excel Table**
```
Create Table (Ctrl+T)
Data Validation → Source: =Table1[Column]
Add/remove rows → dropdown updates
```

**Method 2: OFFSET Formula**
```
Named Range:
=OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),1)
Data Validation → Source: =DynamicList
```

**Method 3: Dynamic Array (365)**
```
Named Range:
=FILTER(Products, Products<>"")
Source: =ProductList
```

---

### Search-able Dropdown

**Large Lists - Find as You Type**

**Excel 365: Use AutoComplete**
```
In-cell dropdown → Type to filter
```

**Older Excel:**
- Use combo box (ActiveX control)
- Or create custom macro

---

### Prevent Duplicates

**Ensure Unique Values**

**Method 1: Custom Formula**
```
Allow: Custom
Formula: =COUNTIF($A$1:$A$100, A1)=1
```

**Method 2: Conditional Check**
```
Allow: Custom
Formula: =COUNTIF($A$1:A1, A1)=1
(Allows first instance only)
```

**Error Message:**
```
Title: Duplicate Entry
Message: This value already exists. Please enter a unique value.
```

---

### Restrict to Named Range Values

**Only Allow Values from List**

**Setup:**
```
Named Range: ValidCodes = {ABC, DEF, GHI}
Validation:
  Allow: Custom
  Formula: =COUNTIF(ValidCodes, A1)>0
```

---

### Combine Multiple Conditions

**AND Logic**
```
Allow: Custom
Formula: =AND(A1>=10, A1<=100, MOD(A1,5)=0)
(Between 10-100 AND multiple of 5)
```

**OR Logic**
```
Allow: Custom
Formula: =OR(A1="Approved", A1="Pending", A1="Rejected")
```

---

### Validate Against Another Sheet

**Reference Different Sheet**
```
Allow: List
Source: =Sheet2!$A$1:$A$100
```

**Or Named Range:**
```
Define Name: Sheet2!$A$1:$A$100 as "RemoteList"
Source: =RemoteList
```

---

## Real-World Examples

### Employee Data Entry Form

**Employee ID:**
```
Allow: Custom
Formula: =AND(LEN(A2)=6, ISNUMBER(--A2))
Message: Enter 6-digit employee ID
```

**Department:**
```
Allow: List
Source: HR,IT,Sales,Marketing,Finance
```

**Start Date:**
```
Allow: Date
Data: greater than or equal to
Start date: 1/1/2000
```

**Salary:**
```
Allow: Decimal
Data: between
Minimum: 30000
Maximum: 200000
```

---

### Inventory Management

**Product Code:**
```
Allow: List
Source: =Products[Code]
(From Table)
```

**Quantity:**
```
Allow: Whole number
Data: greater than or equal to
Minimum: 0
```

**Reorder Level:**
```
Allow: Custom
Formula: =B2<C2
(Quantity < Reorder Point)
Message: Warning: Below reorder level!
```

---

### Survey Form

**Age Group:**
```
Allow: List
Source: 18-24,25-34,35-44,45-54,55-64,65+
```

**Rating (1-5):**
```
Allow: Whole number
Data: between
Minimum: 1
Maximum: 5
```

**Email:**
```
Allow: Custom
Formula: =AND(ISNUMBER(FIND("@",A2)), ISNUMBER(FIND(".",A2)))
```

---

### Project Tracker

**Status:**
```
Allow: List
Source: Not Started,In Progress,Complete,On Hold,Cancelled
```

**Due Date:**
```
Allow: Date
Data: greater than
Date: =TODAY()
Message: Due date must be in future
```

**Priority:**
```
Allow: List
Source: Low,Medium,High,Critical
```

---

## Best Practices

### Design Guidelines

**User Experience:**
- ✅ Always provide input messages
- ✅ Clear, helpful error messages
- ✅ Use dropdowns when options limited
- ✅ Test validation thoroughly
- ✅ Consider user skill level

**Error Messages:**
```
❌ "Invalid entry"
✅ "Please enter a number between 1 and 100"

❌ "Error"
✅ "Email must contain @ and . symbols"
```

**Input Messages:**
```
❌ "Enter data"
✅ "Select product from dropdown or type to search"
```

### Data Quality

**Prevent Errors:**
- ✅ Validate at entry (not after)
- ✅ Restrict to valid options
- ✅ Use formulas for complex rules
- ✅ Prevent duplicates where needed
- ✅ Require critical fields

**Maintain Flexibility:**
- ⚠️ Don't over-restrict
- ⚠️ Allow exceptions when reasonable
- ⚠️ Use Warning vs Stop appropriately

### Performance

**Large Datasets:**
- ✅ Limit validation ranges
- ✅ Use tables for dynamic lists
- ✅ Avoid complex formulas if possible
- ✅ Reference static lists when appropriate

### Maintenance

**Documentation:**
- ✅ Document validation rules
- ✅ Keep lists updated
- ✅ Test after changes
- ✅ Use named ranges (easier to update)

---

## Troubleshooting

### Common Issues

**Dropdown Not Showing:**
```
Check: "In-cell dropdown" is checked
Check: Source range exists
Check: No extra spaces in list
```

**Validation Not Working:**
```
Check: Cells don't have existing validation
Check: Formula returns TRUE/FALSE
Check: Absolute references where needed ($)
```

**INDIRECT Not Working:**
```
Check: Named range exists
Check: Spelling matches exactly
Check: Named range has no spaces
```

**Can't Paste into Validated Cells:**
```
Solution: Temporarily remove validation
Or: Use Paste Special → Values
```

---

## Quick Reference

| Type | Use Case | Example |
|------|----------|---------|
| List | Limited options | Dropdown menu |
| Whole Number | Integer values | Age, Quantity |
| Decimal | Numbers with decimals | Price, Percentage |
| Date | Date ranges | Due dates |
| Time | Time ranges | Business hours |
| Text Length | Character limits | Passwords, Codes |
| Custom | Complex rules | No duplicates, Email |

---

**[⬆ Back to Main README](../../README.md)**
