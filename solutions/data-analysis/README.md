# Data Analysis Solutions

> **Statistical analysis, trends, distributions, and insights**

## Quick Solutions

| I want to... | Solution |
|--------------|----------|
| Find average/median | [Central Tendency](#central-tendency) |
| Measure spread/variance | [Dispersion](#dispersion-measures) |
| Find percentiles/quartiles | [Percentiles](#percentiles-and-quartiles) |
| Rank data | [Ranking](#ranking-data) |
| Find outliers | [Outlier Detection](#outlier-detection) |
| Analyze frequency | [Frequency Distribution](#frequency-distribution) |
| Calculate correlation | [Correlation](#correlation-analysis) |
| Trend analysis | [Trend Lines](#trend-analysis) |
| Compare groups | [Group Comparisons](#group-comparisons) |
| Forecast future values | [Forecasting](#forecasting) |

---

## Central Tendency

### Mean (Average)
```excel
=AVERAGE(A1:A100)              → Arithmetic mean
=AVERAGEIF(A:A, ">0", B:B)     → Average with criteria
=AVERAGEIFS(C:C, A:A, "North", B:B, ">100")  → Multiple criteria
```

### Median (Middle Value)
```excel
=MEDIAN(A1:A100)               → Middle value (50th percentile)
```

### Mode (Most Frequent)
```excel
=MODE.SNGL(A1:A100)            → Single most common value
=MODE.MULT(A1:A100)            → All modes (array, Excel 365)
```

### Trimmed Mean (Exclude Extremes)
```excel
=TRIMMEAN(A1:A100, 0.1)        → Exclude top/bottom 5% each
```

### Weighted Average
```excel
=SUMPRODUCT(Values, Weights) / SUM(Weights)
=SUMPRODUCT(A1:A10, B1:B10) / SUM(B1:B10)
```

### When to Use Each

| Measure | Best For |
|---------|----------|
| Mean | Symmetric data, no outliers |
| Median | Skewed data, with outliers |
| Mode | Categorical data |
| Trimmed Mean | Data with extreme values |

---

## Dispersion Measures

### Standard Deviation
```excel
=STDEV.S(A1:A100)              → Sample standard deviation
=STDEV.P(A1:A100)              → Population standard deviation
```

### Variance
```excel
=VAR.S(A1:A100)                → Sample variance
=VAR.P(A1:A100)                → Population variance
```

### Range
```excel
=MAX(A1:A100) - MIN(A1:A100)
```

### Interquartile Range (IQR)
```excel
=QUARTILE.INC(A1:A100, 3) - QUARTILE.INC(A1:A100, 1)
```

### Coefficient of Variation
```excel
=STDEV.S(A1:A100) / AVERAGE(A1:A100)
```
(Relative variability - useful for comparing datasets)

### Mean Absolute Deviation
```excel
=AVERAGE(ABS(A1:A100 - AVERAGE(A1:A100)))
```
(Enter as array formula or use AVERAGEIF approach)

---

## Percentiles and Quartiles

### Specific Percentile
```excel
=PERCENTILE.INC(A1:A100, 0.9)     → 90th percentile
=PERCENTILE.INC(A1:A100, 0.25)    → 25th percentile (Q1)
=PERCENTILE.INC(A1:A100, 0.5)     → 50th percentile (Median)
=PERCENTILE.INC(A1:A100, 0.75)    → 75th percentile (Q3)
```

### Quartiles
```excel
=QUARTILE.INC(A1:A100, 0)    → Minimum
=QUARTILE.INC(A1:A100, 1)    → Q1 (25th percentile)
=QUARTILE.INC(A1:A100, 2)    → Q2 (Median)
=QUARTILE.INC(A1:A100, 3)    → Q3 (75th percentile)
=QUARTILE.INC(A1:A100, 4)    → Maximum
```

### Five-Number Summary
| Statistic | Formula |
|-----------|---------|
| Min | `=MIN(A:A)` |
| Q1 | `=QUARTILE.INC(A:A, 1)` |
| Median | `=MEDIAN(A:A)` |
| Q3 | `=QUARTILE.INC(A:A, 3)` |
| Max | `=MAX(A:A)` |

### Percentile Rank
What percentile is a specific value?
```excel
=PERCENTRANK.INC(A1:A100, value)
=PERCENTRANK.INC(A1:A100, 75)     → 0.65 means 65th percentile
```

---

## Ranking Data

### Simple Rank
```excel
=RANK.EQ(A1, $A$1:$A$100, 0)      → Descending (largest = 1)
=RANK.EQ(A1, $A$1:$A$100, 1)      → Ascending (smallest = 1)
```

### Average Rank for Ties
```excel
=RANK.AVG(A1, $A$1:$A$100, 0)     → Ties get average rank
```

### Top N Values
```excel
=LARGE(A1:A100, 1)               → Largest value
=LARGE(A1:A100, 5)               → 5th largest
=SMALL(A1:A100, 1)               → Smallest value
=SMALL(A1:A100, 5)               → 5th smallest
```

### Dynamic Top 5 List (Excel 365)
```excel
=SORT(A1:B100, 2, -1)            → Sort by column 2, descending
=TAKE(SORT(A1:B100, 2, -1), 5)   → Top 5 rows
```

### Rank with Criteria
Rank within groups:
```excel
=SUMPRODUCT((B:B=B1)*(C:C>C1))+1
```
(Counts how many in same group have higher value, +1)

---

## Outlier Detection

### IQR Method (Most Common)
```excel
Q1:          =QUARTILE.INC(A:A, 1)
Q3:          =QUARTILE.INC(A:A, 3)
IQR:         =Q3 - Q1
Lower Bound: =Q1 - 1.5 * IQR
Upper Bound: =Q3 + 1.5 * IQR
Is Outlier:  =OR(A1 < LowerBound, A1 > UpperBound)
```

### Z-Score Method
```excel
Z-Score:     =(A1 - AVERAGE($A:$A)) / STDEV.S($A:$A)
Is Outlier:  =ABS(Z-Score) > 3    → TRUE if outlier
```

### Flag Outliers
```excel
=IF(OR(A1 < $B$1, A1 > $B$2), "Outlier", "Normal")
```
Where B1 = lower bound, B2 = upper bound

### Count Outliers
```excel
=COUNTIF(A:A, "<" & LowerBound) + COUNTIF(A:A, ">" & UpperBound)
```

---

## Frequency Distribution

### FREQUENCY Function
```excel
=FREQUENCY(data_array, bins_array)
```

### Full Example

**Data in A1:A100, Bins in C1:C5:**
| Bin | Meaning |
|-----|---------|
| 20 | ≤20 |
| 40 | 21-40 |
| 60 | 41-60 |
| 80 | 61-80 |
| 100 | 81-100 |

**Formula:** `=FREQUENCY(A1:A100, C1:C5)`

**Result (in D1:D6):**
| Count | Range |
|-------|-------|
| 12 | ≤20 |
| 25 | 21-40 |
| 35 | 41-60 |
| 20 | 61-80 |
| 8 | 81-100 |
| 0 | >100 |

### Using COUNTIFS for Bins
```excel
=COUNTIFS(A:A, ">="&C1, A:A, "<"&C2)
```

### Histogram
1. Select frequency results
2. Insert → Charts → Histogram
3. Or use Bar chart with bins as labels

---

## Correlation Analysis

### Correlation Coefficient
```excel
=CORREL(A1:A100, B1:B100)
```

**Interpretation:**
| Value | Meaning |
|-------|---------|
| +1 | Perfect positive |
| +0.7 to +1 | Strong positive |
| +0.3 to +0.7 | Moderate positive |
| -0.3 to +0.3 | Weak/None |
| -0.7 to -0.3 | Moderate negative |
| -1 to -0.7 | Strong negative |
| -1 | Perfect negative |

### Covariance
```excel
=COVARIANCE.S(A1:A100, B1:B100)    → Sample
=COVARIANCE.P(A1:A100, B1:B100)    → Population
```

### R-Squared (Coefficient of Determination)
```excel
=RSQ(known_ys, known_xs)
=CORREL(A:A, B:B)^2                → Square of correlation
```

---

## Trend Analysis

### Linear Trend Line Values
```excel
=TREND(known_ys, known_xs, new_xs)
```

### Slope of Trend
```excel
=SLOPE(known_ys, known_xs)
```
(Rate of change per unit X)

### Y-Intercept
```excel
=INTERCEPT(known_ys, known_xs)
```

### Full Trend Equation
```
Y = SLOPE * X + INTERCEPT
```

### Example - Sales Trend

| Month | Sales |
|-------|-------|
| 1 | 100 |
| 2 | 115 |
| 3 | 125 |
| 4 | 140 |

```excel
Slope:      =SLOPE(B1:B4, A1:A4)     → 13.5
Intercept:  =INTERCEPT(B1:B4, A1:A4) → 85.5
Month 5:    =TREND(B1:B4, A1:A4, 5)  → 153
```

### Growth Rate
```excel
=GROWTH(known_ys, known_xs, new_xs)   → Exponential trend
=LOGEST(known_ys, known_xs)            → Growth rate
```

---

## Group Comparisons

### Compare Averages
```excel
Group A Average:  =AVERAGEIF(Category, "A", Values)
Group B Average:  =AVERAGEIF(Category, "B", Values)
Difference:       =GroupA - GroupB
% Difference:     =(GroupA - GroupB) / GroupB
```

### Summary by Group

| Group | Count | Sum | Average | Min | Max |
|-------|-------|-----|---------|-----|-----|
| A | =COUNTIF | =SUMIF | =AVERAGEIF | - | - |
| B | =COUNTIF | =SUMIF | =AVERAGEIF | - | - |

### Min/Max by Group
```excel
=MINIFS(Values, Category, "A")
=MAXIFS(Values, Category, "A")
```

### Full Comparison Table
```excel
Count:   =COUNTIF(Cat, A1)
Sum:     =SUMIF(Cat, A1, Values)
Average: =AVERAGEIF(Cat, A1, Values)
StdDev:  =AGGREGATE(8, 5, IF(Cat=A1, Values))
```

---

## Forecasting

### Simple Linear Forecast
```excel
=FORECAST.LINEAR(new_x, known_ys, known_xs)
```

### FORECAST.ETS (Seasonal/Time Series)
```excel
=FORECAST.ETS(target_date, values, dates)
```

**Options:**
```excel
=FORECAST.ETS(target, values, dates, seasonality, data_completion, aggregation)
```

### Moving Average
```excel
3-Period: =AVERAGE(A1:A3)      → Copy down
5-Period: =AVERAGE(A1:A5)
```

### Exponential Smoothing
```excel
=Previous_Smooth + Alpha * (Actual - Previous_Smooth)
```
Where Alpha = smoothing factor (0.1 to 0.3 typical)

### Forecast Sheet (Quick Method)
1. Select date + value columns
2. Data → Forecast Sheet
3. Configure options
4. Create

---

## Statistical Tests Quick Reference

### T-Test (Compare Two Means)
```excel
=T.TEST(array1, array2, tails, type)
```
- tails: 1 or 2
- type: 1=paired, 2=equal variance, 3=unequal variance

### Chi-Square Test
```excel
=CHISQ.TEST(actual_range, expected_range)
```

### F-Test (Compare Variances)
```excel
=F.TEST(array1, array2)
```

---

## Data Analysis Toolpak

Enable: File → Options → Add-ins → Analysis ToolPak

### Available Tools
| Tool | Purpose |
|------|---------|
| Descriptive Statistics | Summary stats |
| Histogram | Frequency distribution |
| Correlation | Correlation matrix |
| Regression | Linear regression |
| t-Test | Compare means |
| ANOVA | Compare multiple groups |
| Moving Average | Smoothed trend |
| Exponential Smoothing | Forecast |

### Using Descriptive Statistics
1. Data → Data Analysis → Descriptive Statistics
2. Select input range
3. Check "Summary statistics"
4. Click OK

**Output includes:** Mean, StdError, Median, Mode, StdDev, Variance, Kurtosis, Skewness, Range, Min, Max, Sum, Count

---

## Common Formulas Summary

| Analysis | Formula |
|----------|---------|
| Average | `=AVERAGE(range)` |
| Median | `=MEDIAN(range)` |
| Std Dev | `=STDEV.S(range)` |
| Variance | `=VAR.S(range)` |
| Min | `=MIN(range)` |
| Max | `=MAX(range)` |
| Range | `=MAX(range)-MIN(range)` |
| Percentile | `=PERCENTILE.INC(range, k)` |
| Rank | `=RANK.EQ(value, range)` |
| Correlation | `=CORREL(range1, range2)` |
| Slope | `=SLOPE(y_range, x_range)` |
| Forecast | `=FORECAST.LINEAR(x, y_range, x_range)` |

---

## Related Solutions

- [Conditional Calculations](../conditional-calculations/README.md) - SUMIFS, COUNTIFS
- [Reporting](../reporting/README.md) - Display analysis results
- [Lookups](../lookups/README.md) - Retrieve related data

---

[🏠 Back to Home](../../README.md) | [🎯 All Solutions](../README.md)
