# Goal-Based Solutions Library

> **Find solutions by what you want to accomplish - your roadmap to Excel mastery**

## How to Use This Library

**Don't know the function name?** No problem! Find what you want to DO, and we'll show you how.

---

## Solution Categories

### 📊 [Data Cleaning & Transformation](./data-cleaning/README.md)
Clean up messy data, fix formatting, standardize entries

| I want to... | Go here |
|--------------|---------|
| Remove extra spaces | [Data Cleaning](./data-cleaning/README.md#remove-extra-spaces) |
| Fix capitalization | [Data Cleaning](./data-cleaning/README.md#fix-capitalization) |
| Split names into columns | [Data Cleaning](./data-cleaning/README.md#split-names) |
| Combine columns | [Data Cleaning](./data-cleaning/README.md#combine-columns) |
| Convert text to numbers | [Data Cleaning](./data-cleaning/README.md#convert-text-to-numbers) |
| Remove duplicates | [Data Cleaning](./data-cleaning/README.md#remove-duplicates) |

---

### 🔍 [Lookups & Data Retrieval](./lookups/README.md)
Find values in tables, match data between sheets

| I want to... | Go here |
|--------------|---------|
| Look up a value | [Lookups](./lookups/README.md#basic-lookup) |
| Look up with multiple criteria | [Lookups](./lookups/README.md#lookup-with-multiple-criteria) |
| Return multiple columns | [Lookups](./lookups/README.md#return-multiple-columns) |
| Get all matches (not just first) | [Lookups](./lookups/README.md#return-all-matches) |
| Handle lookup errors | [Lookups](./lookups/README.md#handle-lookup-errors) |
| Two-way lookup | [Lookups](./lookups/README.md#two-way-lookup) |

---

### 🔢 [Conditional Calculations](./conditional-calculations/README.md)
Sum, count, or average based on criteria

| I want to... | Go here |
|--------------|---------|
| Sum based on one condition | [Conditional Calcs](./conditional-calculations/README.md) |
| Sum based on multiple conditions | [Conditional Calcs](./conditional-calculations/README.md) |
| Count items matching criteria | [Conditional Calcs](./conditional-calculations/README.md) |
| Average with conditions | [Conditional Calcs](./conditional-calculations/README.md) |
| Sum across multiple sheets | [Conditional Calcs](./conditional-calculations/README.md) |

---

### 📅 [Date & Time Calculations](./dates-times/README.md)
Work with dates, calculate durations, find business days

| I want to... | Go here |
|--------------|---------|
| Calculate age from birthdate | [Dates & Times](./dates-times/README.md) |
| Find business days between dates | [Dates & Times](./dates-times/README.md) |
| Add/subtract months | [Dates & Times](./dates-times/README.md) |
| Get last day of month | [Dates & Times](./dates-times/README.md) |
| Calculate hours worked | [Dates & Times](./dates-times/README.md) |

---

### 📝 [Text Manipulation](./text-manipulation/README.md)
Extract, combine, and transform text

| I want to... | Go here |
|--------------|---------|
| Extract first/last name | [Text Manipulation](./text-manipulation/README.md) |
| Get numbers from text | [Text Manipulation](./text-manipulation/README.md) |
| Format phone numbers | [Text Manipulation](./text-manipulation/README.md) |
| Create email from names | [Text Manipulation](./text-manipulation/README.md) |
| Count words in a cell | [Text Manipulation](./text-manipulation/README.md) |

---

### 💰 [Financial Calculations](./financial/README.md)
Loans, investments, depreciation

| I want to... | Go here |
|--------------|---------|
| Calculate loan payments | [Financial](./financial/README.md) |
| Create amortization schedule | [Financial](./financial/README.md) |
| Find investment returns | [Financial](./financial/README.md) |
| Calculate depreciation | [Financial](./financial/README.md) |
| Compound interest | [Financial](./financial/README.md) |

---

### 📈 [Data Analysis](./data-analysis/README.md)
Rankings, trends, statistics

| I want to... | Go here |
|--------------|---------|
| Find top/bottom N values | [Data Analysis](./data-analysis/README.md) |
| Rank values | [Data Analysis](./data-analysis/README.md) |
| Calculate running totals | [Data Analysis](./data-analysis/README.md) |
| Find outliers | [Data Analysis](./data-analysis/README.md) |
| Year-over-year comparison | [Data Analysis](./data-analysis/README.md) |

---

### 📊 [Reporting & Dashboards](./reporting/README.md)
Dynamic reports, formatting, visualization

| I want to... | Go here |
|--------------|---------|
| Create dynamic named ranges | [Reporting](./reporting/README.md) |
| Make in-cell charts | [Reporting](./reporting/README.md) |
| Build KPI scorecards | [Reporting](./reporting/README.md) |
| Create dropdown filters | [Reporting](./reporting/README.md) |
| Format numbers dynamically | [Reporting](./reporting/README.md) |

---

### ⚠️ [Error Handling](./error-handling/README.md)
Catch and prevent formula errors

| I want to... | Go here |
|--------------|---------|
| Handle #N/A errors | [Error Handling](./error-handling/README.md#handle-na-errors) |
| Prevent #DIV/0! errors | [Error Handling](./error-handling/README.md#prevent-division-by-zero) |
| Catch any error | [Error Handling](./error-handling/README.md#handle-any-error) |
| Create error-proof formulas | [Error Handling](./error-handling/README.md#create-error-proof-formulas) |

---

### 🚀 [Advanced Techniques](./advanced/README.md)
Array formulas, lambda functions, complex patterns

| I want to... | Go here |
|--------------|---------|
| Use dynamic arrays | [Advanced](./advanced/README.md) |
| Create custom functions | [Advanced](./advanced/README.md) |
| Recursive calculations | [Advanced](./advanced/README.md) |
| Matrix operations | [Advanced](./advanced/README.md) |

---

## Most Common Tasks - Quick Reference

### The Top 10 Things People Need to Do

1. **Look up a value**
   ```excel
   =XLOOKUP(A2, IDs, Names, "Not found")
   ```

2. **Sum with conditions**
   ```excel
   =SUMIFS(Sales, Region, "North", Year, 2024)
   ```

3. **Remove duplicates**
   ```excel
   =UNIQUE(A2:A100)
   ```

4. **Calculate percentage**
   ```excel
   =A2/SUM($A$2:$A$10)
   ```

5. **Find and replace text**
   ```excel
   =SUBSTITUTE(A2, "old", "new")
   ```

6. **Calculate date difference**
   ```excel
   =DATEDIF(A2, B2, "D")    // Days
   =DATEDIF(A2, B2, "Y")    // Years
   ```

7. **Handle errors**
   ```excel
   =IFERROR(formula, "Error message")
   ```

8. **Combine text**
   ```excel
   =TEXTJOIN(", ", TRUE, A2:D2)
   ```

9. **Create running total**
   ```excel
   =SUM($A$2:A2)
   ```

10. **Rank values**
    ```excel
    =RANK.EQ(A2, $A$2:$A$100, 0)
    ```

---

## Can't Find What You Need?

1. **Check the [Functions A-Z](../13-Quick-Reference/Functions-A-Z.md)** - Complete list of all Excel functions
2. **Browse by category** in the [Functions Reference](../functions/README.md)
3. **Search this repository** using GitHub search (press `/`)

---

[🏠 Back to Home](../README.md)
