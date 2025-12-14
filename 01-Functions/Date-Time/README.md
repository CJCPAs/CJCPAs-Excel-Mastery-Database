# 📅 Date & Time Functions

> **25+ functions for date and time calculations, perfect for scheduling, age calculations, and time tracking**

## 📋 Table of Contents

- [Current Date & Time](#current-date--time)
- [Date Creation & Extraction](#date-creation--extraction)
- [Date Calculations](#date-calculations)
- [Time Functions](#time-functions)
- [Working Days & Business Dates](#working-days--business-dates)
- [Date Formatting & Conversion](#date-formatting--conversion)

---

## Current Date & Time

### TODAY
**Returns current date (no time)**

**Syntax:** `=TODAY()`

**Examples:**
```excel
=TODAY()                            → 12/14/2025
=TODAY()+7                          → Date 7 days from now
=TODAY()-30                         → Date 30 days ago
```

**Real-World Uses:**
- Calculate age: `=DATEDIF(Birthdate, TODAY(), "Y")`
- Days until deadline: `=Deadline-TODAY()`
- Overdue items: `=IF(DueDate<TODAY(), "Overdue", "OK")`
- Current month: `=TEXT(TODAY(), "MMMM")`

**Important:** Volatile function - recalculates every time worksheet changes

---

### NOW
**Returns current date and time**

**Syntax:** `=NOW()`

**Examples:**
```excel
=NOW()                              → 12/14/2025 2:33 PM
=INT(NOW())                         → Just the date portion
=MOD(NOW(), 1)                      → Just the time portion
=NOW()+1                            → Same time tomorrow
```

**Real-World Uses:**
- Timestamp entries
- Calculate hours worked: `=NOW()-ClockIn`
- Time until event: `=EventTime-NOW()`

**Important:** Also volatile - updates continuously

---

## Date Creation & Extraction

### DATE
**Creates date from year, month, day**

**Syntax:** `=DATE(year, month, day)`

**Examples:**
```excel
=DATE(2025, 12, 14)                 → 12/14/2025
=DATE(A1, B1, C1)                   → Date from separate cells
=DATE(2025, 13, 1)                  → 1/1/2026 (auto-adjusts)
=DATE(2025, 1, 32)                  → 2/1/2025 (auto-adjusts)
```

**Real-World Uses:**
- Build dates from parts
- First day of month: `=DATE(YEAR(A1), MONTH(A1), 1)`
- End of year: `=DATE(YEAR(A1), 12, 31)`
- Next birthday: `=DATE(YEAR(TODAY()), MONTH(Birthdate), DAY(Birthdate))`

**Auto-Adjustment Examples:**
```excel
=DATE(2025, 0, 1)                   → 12/1/2024 (month 0 = previous Dec)
=DATE(2025, 3, 0)                   → 2/28/2025 (day 0 = last day prev month)
```

---

### YEAR
**Extracts year from date**

**Syntax:** `=YEAR(serial_number)`

**Examples:**
```excel
=YEAR(TODAY())                      → 2025
=YEAR("12/14/2025")                 → 2025
=YEAR(A1)                           → Extract year from date
```

**Real-World Uses:**
- Filter by year: `=YEAR(Date)=2025`
- Age calculation: `=YEAR(TODAY())-YEAR(Birthdate)`
- Fiscal year: `=IF(MONTH(Date)>=7, YEAR(Date)+1, YEAR(Date))`

---

### MONTH
**Extracts month from date (1-12)**

**Syntax:** `=MONTH(serial_number)`

**Examples:**
```excel
=MONTH(TODAY())                     → 12
=MONTH("12/14/2025")                → 12
=TEXT(A1, "MMMM")                   → "December" (month name)
```

**Real-World Uses:**
- Group by month
- Quarter calculation: `=ROUNDUP(MONTH(Date)/3, 0)`
- Month name: `=TEXT(DATE(2025, A1, 1), "MMMM")`

---

### DAY
**Extracts day from date (1-31)**

**Syntax:** `=DAY(serial_number)`

**Examples:**
```excel
=DAY(TODAY())                       → 14
=DAY("12/14/2025")                  → 14
```

---

### WEEKDAY
**Returns day of week (1-7)**

**Syntax:** `=WEEKDAY(serial_number, [return_type])`

**Return Types:**
- **1** (default): Sunday=1, Monday=2, ..., Saturday=7
- **2**: Monday=1, Tuesday=2, ..., Sunday=7
- **3**: Monday=0, Tuesday=1, ..., Sunday=6

**Examples:**
```excel
=WEEKDAY(TODAY())                   → 7 (if today is Saturday)
=WEEKDAY(A1, 2)                     → Monday-based numbering
=TEXT(A1, "DDDD")                   → "Saturday" (day name)
```

**Real-World Uses:**
- Check if weekend: `=WEEKDAY(Date, 2)>5`
- Highlight weekends: `=OR(WEEKDAY(Date)=1, WEEKDAY(Date)=7)`
- Count Mondays: `=SUMPRODUCT((WEEKDAY(Dates, 2)=1)*1)`

---

### WEEKNUM
**Returns week number of year (1-53)**

**Syntax:** `=WEEKNUM(serial_number, [return_type])`

**Examples:**
```excel
=WEEKNUM(TODAY())                   → Week number (e.g., 50)
=WEEKNUM(A1, 2)                     → Week starting Monday
```

**Use Cases:**
- Weekly reports
- Group sales by week
- Production planning

---

### ISOWEEKNUM
**ISO week number (weeks start Monday)**

**Syntax:** `=ISOWEEKNUM(date)`

**Example:**
```excel
=ISOWEEKNUM(TODAY())                → ISO week number
```

---

## Date Calculations

### DATEDIF
**Calculate difference between dates**

**Syntax:** `=DATEDIF(start_date, end_date, unit)`

**Units:**
- **"Y"**: Complete years
- **"M"**: Complete months
- **"D"**: Days
- **"MD"**: Days ignoring months and years
- **"YM"**: Months ignoring years
- **"YD"**: Days ignoring years

**Examples:**
```excel
=DATEDIF(A1, TODAY(), "Y")          → Age in years
=DATEDIF(StartDate, EndDate, "M")   → Months between
=DATEDIF(StartDate, EndDate, "D")   → Days between
```

**Age in Years and Months:**
```excel
=DATEDIF(Birthdate, TODAY(), "Y") & " years, " & DATEDIF(Birthdate, TODAY(), "YM") & " months"
```

**Real-World Uses:**
- Employee tenure: `=DATEDIF(HireDate, TODAY(), "Y")`
- Project duration: `=DATEDIF(StartDate, EndDate, "D")`
- Subscription length: `=DATEDIF(StartDate, TODAY(), "M")`

**Note:** DATEDIF is undocumented but widely used and reliable

---

### EDATE
**Add or subtract months from date**

**Syntax:** `=EDATE(start_date, months)`

**Examples:**
```excel
=EDATE(TODAY(), 1)                  → Date 1 month from today
=EDATE(TODAY(), -3)                 → Date 3 months ago
=EDATE(A1, 12)                      → Date 1 year from A1
```

**Real-World Uses:**
- Subscription renewal: `=EDATE(StartDate, 12)`
- Payment schedule: `=EDATE(FirstPayment, Period*1)`
- Project milestones: `=EDATE(ProjectStart, 3)`

**Important:** Maintains day of month (or adjusts to last day)
```excel
=EDATE("1/31/2025", 1)              → 2/28/2025 (Feb has fewer days)
```

---

### EOMONTH
**Last day of month (offset by months)**

**Syntax:** `=EOMONTH(start_date, months)`

**Examples:**
```excel
=EOMONTH(TODAY(), 0)                → Last day of current month
=EOMONTH(TODAY(), 1)                → Last day of next month
=EOMONTH(TODAY(), -1)               → Last day of previous month
```

**Real-World Uses:**
- Month-end dates for reporting
- Payment due dates (end of month)
- Contract expiration dates

**Get First Day of Month:**
```excel
=EOMONTH(A1, -1) + 1                → First day of month
=DATE(YEAR(A1), MONTH(A1), 1)       → Alternative
```

**Get First Day of Next Month:**
```excel
=EOMONTH(A1, 0) + 1
```

---

### YEARFRAC
**Fraction of year between two dates**

**Syntax:** `=YEARFRAC(start_date, end_date, [basis])`

**Basis:**
- **0** (default): US (NASD) 30/360
- **1**: Actual/actual
- **2**: Actual/360
- **3**: Actual/365
- **4**: European 30/360

**Examples:**
```excel
=YEARFRAC("1/1/2025", "12/31/2025", 1)  → 1.0 (full year)
=YEARFRAC(StartDate, EndDate, 1)         → Fraction of year
```

**Real-World Uses:**
- Interest calculations
- Partial year employment
- Prorated charges

---

### Days Between (Simple Subtraction)
```excel
=EndDate - StartDate                → Days between dates
=TODAY() - Birthdate                → Days since birth
=Deadline - TODAY()                 → Days until deadline
```

**Weeks Between:**
```excel
=(EndDate - StartDate) / 7
```

**Months Between (approximate):**
```excel
=(EndDate - StartDate) / 30.44
```

---

## Time Functions

### TIME
**Creates time from hours, minutes, seconds**

**Syntax:** `=TIME(hour, minute, second)`

**Examples:**
```excel
=TIME(14, 30, 0)                    → 2:30 PM
=TIME(9, 0, 0)                      → 9:00 AM
=TIME(A1, B1, C1)                   → Time from cells
```

**Auto-Adjustment:**
```excel
=TIME(25, 0, 0)                     → 1:00 AM (25 hours = 1 day + 1 hour)
=TIME(0, 90, 0)                     → 1:30 AM (90 min = 1.5 hours)
```

---

### HOUR, MINUTE, SECOND
**Extract time components**

**Syntax:**
```excel
=HOUR(serial_number)                → 0-23
=MINUTE(serial_number)              → 0-59
=SECOND(serial_number)              → 0-59
```

**Examples:**
```excel
=HOUR(NOW())                        → 14 (if 2:33 PM)
=MINUTE(NOW())                      → 33
=SECOND(NOW())                      → 45
```

**Real-World Uses:**
- Extract time parts
- Round to hour: `=TIME(HOUR(A1), 0, 0)`
- Check if business hours: `=AND(HOUR(A1)>=9, HOUR(A1)<17)`

---

### Time Calculations

**Add Hours:**
```excel
=A1 + TIME(2, 0, 0)                 → Add 2 hours
=A1 + 2/24                          → Add 2 hours (alternative)
```

**Add Minutes:**
```excel
=A1 + TIME(0, 30, 0)                → Add 30 minutes
=A1 + 30/1440                       → Add 30 minutes (alternative)
```

**Hours Between Times:**
```excel
=(EndTime - StartTime) * 24
=HOUR(EndTime - StartTime)          → Whole hours only
```

**Minutes Between Times:**
```excel
=(EndTime - StartTime) * 1440
```

**Format as Time:**
```excel
=TEXT(A1, "h:mm AM/PM")             → "2:33 PM"
=TEXT(A1, "HH:mm:ss")               → "14:33:45"
```

---

## Working Days & Business Dates

### NETWORKDAYS
**Number of working days between dates (excludes weekends)**

**Syntax:** `=NETWORKDAYS(start_date, end_date, [holidays])`

**Examples:**
```excel
=NETWORKDAYS("1/1/2025", "12/31/2025")      → Business days in 2025
=NETWORKDAYS(StartDate, EndDate)            → Working days
=NETWORKDAYS(StartDate, EndDate, Holidays)  → Excluding holidays
```

**Real-World Uses:**
- Project timelines
- Billable days
- Work day calculations

**Holidays List:**
Create range with holiday dates, reference in formula:
```excel
=NETWORKDAYS(A1, B1, Holidays!A:A)
```

---

### NETWORKDAYS.INTL
**Working days with custom weekends**

**Syntax:** `=NETWORKDAYS.INTL(start_date, end_date, [weekend], [holidays])`

**Weekend Codes:**
- **1**: Saturday, Sunday (default)
- **2**: Sunday, Monday
- **3**: Monday, Tuesday
- **7**: Friday, Saturday
- **11**: Sunday only
- **12**: Monday only
- **"0000011"**: Custom (1=weekend, 0=workday, starting Monday)

**Examples:**
```excel
=NETWORKDAYS.INTL(A1, B1, 7)                → Fri-Sat weekend
=NETWORKDAYS.INTL(A1, B1, 11)               → Sunday only weekend
=NETWORKDAYS.INTL(A1, B1, "0000011")        → Sat-Sun weekend (custom)
```

**Use Cases:**
- International projects
- Custom work schedules
- 6-day work weeks

---

### WORKDAY
**Date that is N working days from start**

**Syntax:** `=WORKDAY(start_date, days, [holidays])`

**Examples:**
```excel
=WORKDAY(TODAY(), 10)               → Date 10 business days from now
=WORKDAY(StartDate, -5)             → 5 business days before start
=WORKDAY(TODAY(), 20, Holidays)     → Excluding holidays
```

**Real-World Uses:**
- Project deadlines
- Delivery dates
- Response time SLAs

---

### WORKDAY.INTL
**Working day with custom weekends**

**Syntax:** `=WORKDAY.INTL(start_date, days, [weekend], [holidays])`

**Examples:**
```excel
=WORKDAY.INTL(TODAY(), 10, 7)       → 10 workdays (Fri-Sat weekend)
=WORKDAY.INTL(A1, 15, "0000011", Holidays)  → Custom weekend
```

---

## Date Formatting & Conversion

### TEXT (for Dates)
**Format date as text with custom format**

**Syntax:** `=TEXT(value, format_text)`

**Common Date Formats:**
```excel
=TEXT(TODAY(), "MM/DD/YYYY")        → "12/14/2025"
=TEXT(TODAY(), "DD-MMM-YYYY")       → "14-Dec-2025"
=TEXT(TODAY(), "MMMM D, YYYY")      → "December 14, 2025"
=TEXT(TODAY(), "DDDD")              → "Saturday"
=TEXT(TODAY(), "DDD")               → "Sat"
=TEXT(TODAY(), "MMMM")              → "December"
=TEXT(TODAY(), "MMM")               → "Dec"
```

**Date & Time:**
```excel
=TEXT(NOW(), "MM/DD/YYYY hh:mm AM/PM")  → "12/14/2025 02:33 PM"
=TEXT(NOW(), "YYYY-MM-DD HH:mm:ss")     → "2025-12-14 14:33:45"
```

**Custom:**
```excel
=TEXT(TODAY(), "DDD, MMM DD")       → "Sat, Dec 14"
="Today is " & TEXT(TODAY(), "DDDD, MMMM D")  → "Today is Saturday, December 14"
```

---

### DATEVALUE
**Convert text to date serial number**

**Syntax:** `=DATEVALUE(date_text)`

**Examples:**
```excel
=DATEVALUE("12/14/2025")            → 46079 (serial number)
=DATEVALUE("December 14, 2025")     → 46079
=DATEVALUE("14-Dec-2025")           → 46079
```

**Use:** Convert text dates to real dates for calculations

---

### TIMEVALUE
**Convert text to time serial number**

**Syntax:** `=TIMEVALUE(time_text)`

**Examples:**
```excel
=TIMEVALUE("2:33 PM")               → 0.60625 (fraction of day)
=TIMEVALUE("14:33:00")              → 0.60625
```

---

## Practical Examples & Patterns

### Age Calculator
```excel
=DATEDIF(Birthdate, TODAY(), "Y")   → Years
=DATEDIF(Birthdate, TODAY(), "Y") & " years, " & DATEDIF(Birthdate, TODAY(), "YM") & " months"
```

### Days Until Event
```excel
=EventDate - TODAY()                → Days remaining
=IF(EventDate<TODAY(), "Past", EventDate-TODAY() & " days")
```

### Tenure Calculation
```excel
=DATEDIF(HireDate, TODAY(), "Y") & " years, " & DATEDIF(HireDate, TODAY(), "YM") & " months"
```

### Next Birthday
```excel
=DATE(YEAR(TODAY()), MONTH(Birthdate), DAY(Birthdate))
=IF(DATE(YEAR(TODAY()), MONTH(Birthdate), DAY(Birthdate))<TODAY(), 
    DATE(YEAR(TODAY())+1, MONTH(Birthdate), DAY(Birthdate)),
    DATE(YEAR(TODAY()), MONTH(Birthdate), DAY(Birthdate)))
```

### Age Group
```excel
=IFS(DATEDIF(Birthdate,TODAY(),"Y")<18,"Under 18",
     DATEDIF(Birthdate,TODAY(),"Y")<65,"Adult",
     TRUE,"Senior")
```

### Quarter from Date
```excel
="Q" & ROUNDUP(MONTH(A1)/3, 0)      → "Q4"
=CHOOSE(ROUNDUP(MONTH(A1)/3,0),"Q1","Q2","Q3","Q4")
```

### First/Last Day of Quarter
```excel
// First day of quarter
=DATE(YEAR(A1), (ROUNDUP(MONTH(A1)/3,0)-1)*3+1, 1)

// Last day of quarter
=EOMONTH(DATE(YEAR(A1), ROUNDUP(MONTH(A1)/3,0)*3, 1), 0)
```

### Check if Date is Weekend
```excel
=WEEKDAY(A1, 2) > 5                 → TRUE if Sat or Sun
=OR(WEEKDAY(A1)=1, WEEKDAY(A1)=7)   → Alternative
```

### Business Days in Month
```excel
=NETWORKDAYS(DATE(YEAR(A1),MONTH(A1),1), EOMONTH(A1,0))
```

### Time Sheet Hours
```excel
=(ClockOut - ClockIn) * 24          → Hours worked
=TEXT(ClockOut-ClockIn, "h:mm")     → Formatted time
```

### Overtime Calculation
```excel
=MAX(0, (ClockOut-ClockIn)*24 - 8)  → Overtime hours (>8)
```

---

## Quick Reference

| Function | Purpose | Example |
|----------|---------|---------|
| TODAY | Current date | `=TODAY()` |
| NOW | Current date/time | `=NOW()` |
| DATE | Create date | `=DATE(2025,12,14)` |
| YEAR/MONTH/DAY | Extract parts | `=YEAR(A1)` |
| DATEDIF | Date difference | `=DATEDIF(A1,B1,"Y")` |
| EDATE | Add months | `=EDATE(A1,3)` |
| EOMONTH | End of month | `=EOMONTH(A1,0)` |
| NETWORKDAYS | Business days | `=NETWORKDAYS(A1,B1)` |
| WORKDAY | Future workday | `=WORKDAY(TODAY(),10)` |
| WEEKDAY | Day of week | `=WEEKDAY(A1)` |
| TEXT | Format date | `=TEXT(A1,"MM/DD/YYYY")` |

---

**[⬆ Back to Main README](../../README.md)**
