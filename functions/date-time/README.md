# Date & Time Functions

> **Calculate, format, and manipulate dates and times**

## Function Quick Reference

| Function | Purpose | Example |
|----------|---------|---------|
| **TODAY** | Current date | `=TODAY()` → 12/14/2025 |
| **NOW** | Current date/time | `=NOW()` → 12/14/2025 3:45 PM |
| **DATE** | Create date | `=DATE(2025,12,25)` |
| **TIME** | Create time | `=TIME(14,30,0)` → 2:30 PM |
| **YEAR** | Extract year | `=YEAR(A1)` → 2025 |
| **MONTH** | Extract month | `=MONTH(A1)` → 12 |
| **DAY** | Extract day | `=DAY(A1)` → 14 |
| **HOUR** | Extract hour | `=HOUR(A1)` → 15 |
| **MINUTE** | Extract minute | `=MINUTE(A1)` → 30 |
| **SECOND** | Extract second | `=SECOND(A1)` → 0 |
| **WEEKDAY** | Day of week (1-7) | `=WEEKDAY(A1)` → 7 |
| **WEEKNUM** | Week number | `=WEEKNUM(A1)` → 50 |
| **DATEDIF** | Date difference | `=DATEDIF(A1,B1,"Y")` |
| **DAYS** | Days between | `=DAYS(B1,A1)` → 30 |
| **EDATE** | Add months | `=EDATE(A1,3)` |
| **EOMONTH** | End of month | `=EOMONTH(A1,0)` |
| **NETWORKDAYS** | Work days between | `=NETWORKDAYS(A1,B1)` |
| **WORKDAY** | Add work days | `=WORKDAY(A1,10)` |
| **DATEVALUE** | Text to date | `=DATEVALUE("12/25/2025")` |
| **TIMEVALUE** | Text to time | `=TIMEVALUE("2:30 PM")` |

## Common Solutions

### Calculate Age
```excel
=DATEDIF(BirthDate,TODAY(),"Y")
```

### Add Business Days
```excel
=WORKDAY(StartDate, 10, Holidays)
```

### Get Month Name
```excel
=TEXT(A1,"MMMM")  → "December"
```

### First Day of Month
```excel
=DATE(YEAR(A1),MONTH(A1),1)
```

### Last Day of Month
```excel
=EOMONTH(A1,0)
```

### Days Until Date
```excel
=A1-TODAY()
```

---

[📚 Full Date/Time Solutions](../../solutions/dates-times/) | [🏠 Back to Home](../../README.md)
