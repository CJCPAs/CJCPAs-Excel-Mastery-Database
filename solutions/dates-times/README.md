# Date & Time Calculation Solutions

> **Work with dates, calculate durations, find business days**

## Quick Solutions

| I want to... | Solution |
|--------------|----------|
| Get today's date | [TODAY()](#todays-date) |
| Calculate age | [DATEDIF](#calculate-age) |
| Find business days between dates | [NETWORKDAYS](#business-days-between-dates) |
| Add business days to a date | [WORKDAY](#add-business-days) |
| Get last day of month | [EOMONTH](#last-day-of-month) |
| Add/subtract months | [EDATE](#addsubtract-months) |
| Extract year/month/day | [YEAR, MONTH, DAY](#extract-date-parts) |
| Calculate hours worked | [Time Calculations](#calculate-hours-worked) |

---

## Today's Date

### Quick Answer
```excel
=TODAY()           // Returns current date (updates automatically)
=NOW()             // Returns current date AND time
```

### Example Uses
```excel
=TODAY()-A2                    // Days since date in A2
=TODAY()+30                    // 30 days from today
=YEAR(TODAY())                 // Current year
=IF(A2<TODAY(), "Overdue", "Current")
```

### Warning
TODAY() and NOW() are volatile - they recalculate every time the workbook recalculates, which can slow down large workbooks.

---

## Calculate Age

### The Challenge
Calculate someone's age in years from their birthdate.

### Quick Answer
```excel
=DATEDIF(birthdate, TODAY(), "Y")
```

### Full Example

**Data:**
| A | B |
|---|---|
| Name | Birthdate |
| John | 3/15/1985 |
| Jane | 7/22/1990 |

**Formula:** `=DATEDIF(B2, TODAY(), "Y")`

**Result:** Age in complete years

### DATEDIF Parameters

| Code | Returns |
|------|---------|
| "Y" | Complete years |
| "M" | Complete months |
| "D" | Days |
| "YM" | Months beyond complete years |
| "MD" | Days beyond complete months |
| "YD" | Days beyond complete years |

### Age in Years, Months, Days
```excel
=DATEDIF(A2, TODAY(), "Y") & " years, " &
 DATEDIF(A2, TODAY(), "YM") & " months, " &
 DATEDIF(A2, TODAY(), "MD") & " days"
```
**Result:** "35 years, 8 months, 29 days"

### Alternative (Without DATEDIF)
```excel
=INT((TODAY()-A2)/365.25)
```

---

## Business Days Between Dates

### The Challenge
Calculate working days between two dates (excluding weekends).

### Quick Answer
```excel
=NETWORKDAYS(start_date, end_date, [holidays])
```

### Full Example

**Data:**
| A | B |
|---|---|
| Start | End |
| 1/1/2024 | 1/31/2024 |

**Formula:** `=NETWORKDAYS(A2, B2)`

**Result:** `23` (business days in January 2024)

### With Holidays
```excel
=NETWORKDAYS(A2, B2, Holidays)    // Where Holidays is a range of holiday dates
```

### Custom Weekends
```excel
=NETWORKDAYS.INTL(A2, B2, weekend_code, holidays)
```

| Code | Weekend Days |
|------|--------------|
| 1 | Saturday, Sunday (default) |
| 2 | Sunday, Monday |
| 11 | Sunday only |
| 12 | Monday only |

---

## Add Business Days

### The Challenge
Calculate a date N business days from a start date.

### Quick Answer
```excel
=WORKDAY(start_date, days, [holidays])
```

### Full Example

**Data:**
| A | B |
|---|---|
| Start Date | Days to Add |
| 1/15/2024 | 10 |

**Formula:** `=WORKDAY(A2, B2)`

**Result:** `1/29/2024` (10 business days after Jan 15)

### With Custom Weekends
```excel
=WORKDAY.INTL(A2, B2, weekend_code, holidays)
```

### Subtract Business Days
Use negative number:
```excel
=WORKDAY(A2, -5)    // 5 business days before A2
```

---

## Last Day of Month

### The Challenge
Find the last day of a month (which varies: 28, 29, 30, or 31 days).

### Quick Answer
```excel
=EOMONTH(date, 0)           // Last day of same month
=EOMONTH(date, 1)           // Last day of next month
=EOMONTH(date, -1)          // Last day of previous month
```

### Full Example

**Starting Date:** 3/15/2024

| Formula | Result |
|---------|--------|
| `=EOMONTH(A2, 0)` | 3/31/2024 |
| `=EOMONTH(A2, 1)` | 4/30/2024 |
| `=EOMONTH(A2, -1)` | 2/29/2024 |

### First Day of Month
```excel
=EOMONTH(A2, -1) + 1        // First day of current month
=DATE(YEAR(A2), MONTH(A2), 1)   // Alternative
```

### First Day of Next Month
```excel
=EOMONTH(A2, 0) + 1
```

---

## Add/Subtract Months

### The Challenge
Add or subtract a specific number of months from a date.

### Quick Answer
```excel
=EDATE(date, months)        // months can be negative
```

### Full Example

**Starting Date:** 1/31/2024

| Formula | Result |
|---------|--------|
| `=EDATE(A2, 1)` | 2/29/2024 |
| `=EDATE(A2, 3)` | 4/30/2024 |
| `=EDATE(A2, -2)` | 11/30/2023 |

### Note About Month Ends
If the start date is the last day of a month and the result month has fewer days, EDATE returns the last day of the result month.

---

## Extract Date Parts

### Quick Reference
```excel
=YEAR(A2)          // Returns year (e.g., 2024)
=MONTH(A2)         // Returns month number (1-12)
=DAY(A2)           // Returns day (1-31)
=WEEKDAY(A2)       // Returns day of week (1=Sunday by default)
=WEEKNUM(A2)       // Returns week number of year
```

### Get Month Name
```excel
=TEXT(A2, "MMMM")      // "January"
=TEXT(A2, "MMM")       // "Jan"
```

### Get Day Name
```excel
=TEXT(A2, "DDDD")      // "Monday"
=TEXT(A2, "DDD")       // "Mon"
```

### Quarter of Year
```excel
=ROUNDUP(MONTH(A2)/3, 0)                   // Returns 1, 2, 3, or 4
="Q" & ROUNDUP(MONTH(A2)/3, 0)             // Returns "Q1", "Q2", etc.
```

---

## Calculate Hours Worked

### The Challenge
Calculate time duration between start and end times.

### Quick Answer
```excel
=End_Time - Start_Time
```
Format result as time (h:mm) or as number × 24 for decimal hours.

### Full Example

**Data:**
| A | B | C |
|---|---|---|
| In | Out | Hours |
| 8:30 AM | 5:00 PM | =B2-A2 |
| 9:00 AM | 6:30 PM | =B3-A3 |

**Results:**
| In | Out | Hours |
|---|---|---|
| 8:30 AM | 5:00 PM | 8:30 |
| 9:00 AM | 6:30 PM | 9:30 |

### Decimal Hours
```excel
=(B2-A2)*24           // 8.5 hours instead of 8:30
```

### Handle Overnight Shifts
```excel
=IF(B2<A2, B2+1-A2, B2-A2)     // Add 1 day if end time < start time
=MOD(B2-A2, 1)                  // Alternative
```

### Sum Hours Over 24
If you're adding hours that total more than 24:
1. Format cells as `[h]:mm` (square brackets allow hours > 24)

---

## Create a Date from Parts

### Quick Answer
```excel
=DATE(year, month, day)
```

### Full Example
```excel
=DATE(2024, 12, 25)            // 12/25/2024
=DATE(A2, B2, C2)              // From separate columns
=DATE(YEAR(TODAY()), 12, 31)   // Dec 31 of current year
```

### Create Time from Parts
```excel
=TIME(hour, minute, second)
=TIME(14, 30, 0)               // 2:30 PM
```

---

## Date Differences

### Days Between Dates
```excel
=B2-A2                         // Simple subtraction
=DAYS(B2, A2)                  // DAYS function (Excel 2013+)
```

### Months Between Dates
```excel
=DATEDIF(A2, B2, "M")          // Complete months
=(YEAR(B2)-YEAR(A2))*12 + MONTH(B2)-MONTH(A2)   // Alternative
```

### Years Between Dates
```excel
=DATEDIF(A2, B2, "Y")          // Complete years
```

---

## Common Date Formulas

### Is This Date a Weekend?
```excel
=IF(WEEKDAY(A2, 2)>5, "Weekend", "Weekday")
```
(WEEKDAY with 2 makes Monday=1, Sunday=7)

### Next Monday
```excel
=A2 + 8 - WEEKDAY(A2, 2)
```

### Same Day Last Year
```excel
=DATE(YEAR(A2)-1, MONTH(A2), DAY(A2))
```

### Days Until Event
```excel
=EventDate - TODAY()
```

### Due Date (30 Days from Invoice)
```excel
=InvoiceDate + 30
```

---

## Related Solutions

- [Conditional Calculations](../conditional-calculations/README.md) - Sum/count by date ranges
- [Text Manipulation](../text-manipulation/README.md) - Format dates as text
- [Lookups](../lookups/README.md) - Look up by date

---

[🏠 Back to Home](../../README.md) | [🎯 All Solutions](../README.md)
