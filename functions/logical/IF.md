# IF

## What It Does (Plain English)
Makes a decision: checks if something is true, and returns one value if yes, a different value if no. Like asking "Is it raining? If yes, take an umbrella. If no, wear sunglasses."

## Syntax
```
=IF(logical_test, value_if_true, [value_if_false])
```

## Arguments

| Argument | Required? | Description |
|----------|-----------|-------------|
| logical_test | Yes | A condition that evaluates to TRUE or FALSE |
| value_if_true | Yes | What to return if the condition is TRUE |
| value_if_false | No | What to return if FALSE (default: FALSE or 0) |

## Returns
Either value_if_true or value_if_false, depending on the test result.

## Examples

### Example 1: Pass/Fail Based on Score
**Data:**
| A | B |
|---|---|
| Student | Score |
| Alice | 85 |
| Bob | 58 |
| Carol | 72 |

**Formula:** `=IF(B2>=60, "Pass", "Fail")`

**Results:**
| Student | Score | Result |
|---|---|---|
| Alice | 85 | Pass |
| Bob | 58 | Fail |
| Carol | 72 | Pass |

**Explanation:** If Score ≥ 60, show "Pass", otherwise "Fail".

---

### Example 2: Calculate Bonus Based on Sales
**Data:**
| A | B |
|---|---|
| Salesperson | Sales |
| Dave | 12000 |
| Eve | 8500 |

**Formula:** `=IF(B2>10000, B2*0.1, B2*0.05)`

**Results:**
| Salesperson | Sales | Bonus |
|---|---|---|
| Dave | 12000 | 1200 |
| Eve | 8500 | 425 |

**Explanation:** 10% bonus if Sales > 10000, otherwise 5%.

---

### Example 3: Nested IF for Multiple Conditions
**Grade assignment:**

**Formula:** `=IF(B2>=90, "A", IF(B2>=80, "B", IF(B2>=70, "C", IF(B2>=60, "D", "F"))))`

**Results:**
| Score | Grade |
|---|---|
| 95 | A |
| 82 | B |
| 71 | C |
| 65 | D |
| 45 | F |

**Explanation:** Checks each threshold in order. First TRUE wins.

---

### Example 4: IF with AND (Multiple Conditions Required)
**Data:**
| A | B | C |
|---|---|---|
| Employee | Sales | Years |
| Frank | 15000 | 3 |
| Grace | 12000 | 1 |

**Formula:** `=IF(AND(B2>10000, C2>=2), "Eligible", "Not Eligible")`

**Results:**
| Employee | Sales | Years | Status |
|---|---|---|---|
| Frank | 15000 | 3 | Eligible |
| Grace | 12000 | 1 | Not Eligible |

**Explanation:** Must have BOTH Sales > 10000 AND Years ≥ 2.

---

### Example 5: IF with OR (Any Condition Sufficient)
**Formula:** `=IF(OR(A2="Manager", B2>50000), "Executive", "Staff")`

**Explanation:** If title is Manager OR salary > 50000, classify as Executive.

---

### Example 6: IF to Handle Blank Cells
**Data:**
| A | B |
|---|---|
| Item | Quantity |
| Widget | 50 |
| Gadget | (blank) |

**Formula:** `=IF(B2="", "No data", B2*10)`

**Results:**
| Item | Quantity | Extended |
|---|---|---|
| Widget | 50 | 500 |
| Gadget | | No data |

**Explanation:** Checks if cell is blank before calculating.

---

### Example 7: Return Formula or Value
**Formula:** `=IF(A2>100, A2*1.1, A2)`

**Explanation:** You can return the original value unchanged in one branch.

## Common Errors

| Error | Cause | Fix |
|-------|-------|-----|
| Formula shows as text | Cell formatted as text | Format cell as General, re-enter formula |
| #VALUE! | Comparing incompatible types | Ensure you're comparing numbers to numbers, text to text |
| Unexpected result | Text comparisons are case-insensitive | IF treats "ABC" = "abc" as TRUE |
| Logic errors | Wrong operator | Double-check >, <, >=, <=, =, <> |

## Pro Tips

- **Nested IF limit:** You can nest up to 64 IFs, but beyond 3-4, consider IFS or SWITCH
- **Text in quotes:** Return values that are text must be in quotes: `"Pass"` not `Pass`
- **Omitting value_if_false:** Returns FALSE if omitted. Use `=IF(A1>5, "Yes", "")` for blank
- **Boolean shortcuts:** `=IF(A1, "Yes", "No")` works because non-zero numbers are TRUE
- **Combine with IFERROR:** `=IFERROR(IF(A1>0, B1/A1, ""), "Error")` handles both cases

## Related Functions

| Function | When to Use Instead |
|----------|-------------------|
| [IFS](./IFS.md) | Multiple conditions without nesting (Excel 2019+) |
| [SWITCH](./SWITCH.md) | When checking same value against multiple options |
| [CHOOSE](../lookup-reference/CHOOSE.md) | When selecting from a numbered list |
| [IFERROR](./IFERROR.md) | Specifically to catch errors |
| [IFNA](./IFNA.md) | Specifically to catch #N/A errors |

## Version Notes
- **Available in:** All Excel versions
- **Google Sheets:** ✅ Fully compatible
- **LibreOffice Calc:** ✅ Fully compatible

---
[📘 Back to Logical Functions](./README.md) | [🏠 Back to Home](../../README.md)
