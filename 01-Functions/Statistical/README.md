# 📊 Statistical Functions

> **100+ functions for statistical analysis, forecasting, and data insights**

## 📋 Table of Contents

- [Descriptive Statistics](#descriptive-statistics)
- [Central Tendency](#central-tendency)
- [Dispersion & Variability](#dispersion--variability)
- [Distribution Functions](#distribution-functions)
- [Correlation & Regression](#correlation--regression)
- [Forecasting](#forecasting)
- [Ranking & Percentiles](#ranking--percentiles)

---

## Descriptive Statistics

### AVERAGE
**Calculates arithmetic mean**

**Syntax:** `=AVERAGE(number1, [number2], ...)`

**Examples:**
```excel
=AVERAGE(A1:A10)                    → Mean of range
=AVERAGE(10,20,30,40,50)           → Returns 30
=AVERAGE(A:A)                       → Average of column (ignores text/blanks)
```

**Real-World Uses:**
- Average sales per month
- Mean test scores
- Average response time
- Average customer age

**Related Functions:**
- **AVERAGEA** - Includes text (counts as 0)
- **AVERAGEIF** - Conditional average
- **AVERAGEIFS** - Multiple criteria

---

### AVERAGEIF
**Calculates average for cells meeting criteria**

**Syntax:** `=AVERAGEIF(range, criteria, [average_range])`

**Examples:**
```excel
=AVERAGEIF(A1:A10, ">100")                     → Average of values >100
=AVERAGEIF(B1:B10, "Apple", C1:C10)            → Average C where B="Apple"
=AVERAGEIF(Scores, ">=60")                     → Average passing scores
```

**Real-World Uses:**
- Average sales by product
- Mean score for specific category
- Average price above threshold

---

### AVERAGEIFS
**Calculates average with multiple criteria**

**Syntax:** `=AVERAGEIFS(average_range, criteria_range1, criteria1, ...)`

**Examples:**
```excel
=AVERAGEIFS(Sales, Region, "North", Product, "Apple")
=AVERAGEIFS(D:D, A:A, ">100", B:B, "<500")
```

---

## Central Tendency

### MEDIAN
**Returns middle value in dataset**

**Syntax:** `=MEDIAN(number1, [number2], ...)`

**Examples:**
```excel
=MEDIAN(A1:A10)                     → Middle value
=MEDIAN(1,2,3,4,5)                  → Returns 3
=MEDIAN(1,2,3,4,5,6)                → Returns 3.5 (average of middle two)
```

**Use Cases:**
- Housing prices (less affected by outliers than average)
- Income statistics
- Response times

**MEDIAN vs AVERAGE:**
- MEDIAN: Not affected by extreme values
- AVERAGE: Influenced by all values

**Example:**
```
Values: 10, 20, 30, 40, 1000
AVERAGE: 220
MEDIAN: 30 (more representative)
```

---

### MODE.SNGL
**Returns most frequently occurring value**

**Syntax:** `=MODE.SNGL(number1, [number2], ...)`

**Examples:**
```excel
=MODE.SNGL(A1:A10)                  → Most common value
=MODE.SNGL(1,2,2,3,3,3,4)          → Returns 3
```

**Use Cases:**
- Most common shoe size
- Most frequent purchase amount
- Typical grade

**Note:** Returns #N/A if no value appears more than once

---

### MODE.MULT
**Returns array of most frequent values**

**Syntax:** `=MODE.MULT(number1, [number2], ...)`

**Returns:** All values tied for most frequent (as array)

---

## Dispersion & Variability

### STDEV.S
**Standard deviation of sample**

**Syntax:** `=STDEV.S(number1, [number2], ...)`

**Examples:**
```excel
=STDEV.S(A1:A10)                    → Sample standard deviation
```

**Use Cases:**
- Measure of variability/spread
- Quality control
- Risk assessment

**Low STDEV:** Values close to mean (consistent)
**High STDEV:** Values spread out (variable)

---

### STDEV.P
**Standard deviation of entire population**

**Syntax:** `=STDEV.P(number1, [number2], ...)`

**Use:** When you have complete population, not just sample

---

### VAR.S
**Variance of sample (STDEV squared)**

**Syntax:** `=VAR.S(number1, [number2], ...)`

**Examples:**
```excel
=VAR.S(A1:A10)                      → Sample variance
```

---

### VAR.P
**Variance of population**

**Syntax:** `=VAR.P(number1, [number2], ...)`

---

### DEVSQ
**Sum of squared deviations from mean**

**Syntax:** `=DEVSQ(number1, [number2], ...)`

**Use:** Statistical calculations, regression analysis

---

## Distribution Functions

### NORM.DIST
**Normal distribution**

**Syntax:** `=NORM.DIST(x, mean, standard_dev, cumulative)`

**Parameters:**
- `x`: Value to evaluate
- `mean`: Arithmetic mean
- `standard_dev`: Standard deviation
- `cumulative`: TRUE = cumulative, FALSE = probability density

**Examples:**
```excel
=NORM.DIST(42, 40, 1.5, TRUE)      → Cumulative probability
=NORM.DIST(42, 40, 1.5, FALSE)     → Probability density
```

**Use Cases:**
- Quality control
- Test scores
- Manufacturing tolerances

---

### NORM.INV
**Inverse normal distribution**

**Syntax:** `=NORM.INV(probability, mean, standard_dev)`

**Example:**
```excel
=NORM.INV(0.95, 100, 15)           → Value at 95th percentile
```

**Use:** Find cutoff values for given probability

---

### NORM.S.DIST
**Standard normal distribution (mean=0, stdev=1)**

**Syntax:** `=NORM.S.DIST(z, cumulative)`

**Example:**
```excel
=NORM.S.DIST(1.96, TRUE)           → Returns 0.975 (97.5%)
```

---

### T.DIST
**Student's t-distribution**

**Syntax:** `=T.DIST(x, deg_freedom, cumulative)`

**Use:** Small sample sizes, hypothesis testing

---

### T.DIST.2T
**Two-tailed t-distribution**

**Syntax:** `=T.DIST.2T(x, deg_freedom)`

**Use:** Two-tailed hypothesis tests

---

### CHI.SQ.DIST
**Chi-squared distribution**

**Syntax:** `=CHI.SQ.DIST(x, deg_freedom, cumulative)`

**Use:** Goodness of fit tests

---

### F.DIST
**F probability distribution**

**Syntax:** `=F.DIST(x, deg_freedom1, deg_freedom2, cumulative)`

**Use:** Analysis of variance (ANOVA)

---

## Correlation & Regression

### CORREL
**Correlation coefficient between two datasets**

**Syntax:** `=CORREL(array1, array2)`

**Examples:**
```excel
=CORREL(A1:A10, B1:B10)            → Correlation between A and B
```

**Returns:** Value between -1 and 1
- **1**: Perfect positive correlation
- **0**: No correlation
- **-1**: Perfect negative correlation

**Use Cases:**
- Sales vs advertising spend
- Temperature vs ice cream sales
- Study time vs test scores

---

### PEARSON
**Pearson correlation coefficient (same as CORREL)**

**Syntax:** `=PEARSON(array1, array2)`

---

### COVARIANCE.P
**Population covariance**

**Syntax:** `=COVARIANCE.P(array1, array2)`

**Use:** Measure how two variables change together

---

### SLOPE
**Slope of linear regression line**

**Syntax:** `=SLOPE(known_y's, known_x's)`

**Examples:**
```excel
=SLOPE(B1:B10, A1:A10)             → Slope of best-fit line
```

**Use:** Rate of change, trend analysis

---

### INTERCEPT
**Y-intercept of linear regression**

**Syntax:** `=INTERCEPT(known_y's, known_x's)`

**Use:** Where line crosses y-axis

---

### RSQ
**R-squared (coefficient of determination)**

**Syntax:** `=RSQ(known_y's, known_x's)`

**Returns:** Value 0 to 1
- **1**: Perfect fit
- **0**: No relationship

**Use:** Measure how well regression line fits data

---

### LINEST
**Returns statistics for linear regression (array function)**

**Syntax:** `=LINEST(known_y's, [known_x's], [const], [stats])`

**Returns:** Slope, intercept, and statistics

---

## Forecasting

### FORECAST.LINEAR
**Linear forecast (trend line)**

**Syntax:** `=FORECAST.LINEAR(x, known_y's, known_x's)`

**Examples:**
```excel
=FORECAST.LINEAR(A11, B1:B10, A1:A10)    → Predict value for A11
```

**Use Cases:**
- Sales forecasting
- Demand prediction
- Budget projections

---

### FORECAST.ETS
**Exponential smoothing forecast (Excel 2016+)**

**Syntax:** `=FORECAST.ETS(target_date, values, timeline, [seasonality], [data_completion], [aggregation])`

**Examples:**
```excel
=FORECAST.ETS(A11, B1:B10, A1:A10, 1, 1, 1)
```

**Advantages:**
- Handles seasonality
- Better for time series data
- Automatic trend detection

---

### TREND
**Returns values along linear trend**

**Syntax:** `=TREND(known_y's, [known_x's], [new_x's], [const])`

**Examples:**
```excel
=TREND(B1:B10, A1:A10, A11:A15)    → Forecast 5 values
```

---

### GROWTH
**Exponential growth trend**

**Syntax:** `=GROWTH(known_y's, [known_x's], [new_x's], [const])`

**Use:** When data grows exponentially (compound growth)

---

## Ranking & Percentiles

### RANK.EQ
**Rank of number in list**

**Syntax:** `=RANK.EQ(number, ref, [order])`

**Parameters:**
- `number`: Value to rank
- `ref`: Array of values
- `order`: 0 = descending (default), 1 = ascending

**Examples:**
```excel
=RANK.EQ(A2, $A$2:$A$10, 0)        → Rank high to low
=RANK.EQ(A2, $A$2:$A$10, 1)        → Rank low to high
```

**Use Cases:**
- Sales rankings
- Student rankings
- Performance rankings

**Handling Ties:**
- RANK.EQ: Same rank, skips next
- RANK.AVG: Averages ranks for ties

---

### RANK.AVG
**Average rank for tied values**

**Syntax:** `=RANK.AVG(number, ref, [order])`

**Example:**
```
Values: 100, 95, 95, 90
RANK.EQ: 1, 2, 2, 4 (skips 3)
RANK.AVG: 1, 2.5, 2.5, 4 (averages 2 and 3)
```

---

### PERCENTILE.INC
**Returns kth percentile (inclusive)**

**Syntax:** `=PERCENTILE.INC(array, k)`

**Parameters:**
- `array`: Data array
- `k`: Percentile (0 to 1)

**Examples:**
```excel
=PERCENTILE.INC(A1:A100, 0.75)     → 75th percentile (Q3)
=PERCENTILE.INC(A1:A100, 0.5)      → 50th percentile (median)
=PERCENTILE.INC(A1:A100, 0.95)     → 95th percentile
```

**Use Cases:**
- Test score cutoffs
- Income distributions
- Performance benchmarks

---

### PERCENTILE.EXC
**Returns kth percentile (exclusive)**

**Syntax:** `=PERCENTILE.EXC(array, k)`

**Difference:** Uses different calculation method (Excel 2010+)

---

### PERCENTRANK.INC
**Rank as percentage of dataset**

**Syntax:** `=PERCENTRANK.INC(array, x, [significance])`

**Examples:**
```excel
=PERCENTRANK.INC(A1:A100, 85)      → What percentile is 85?
```

**Returns:** Percentage rank (0 to 1)

---

### QUARTILE.INC
**Returns quartile of dataset**

**Syntax:** `=QUARTILE.INC(array, quart)`

**Quart values:**
- **0**: Minimum
- **1**: Q1 (25th percentile)
- **2**: Q2 (median, 50th percentile)
- **3**: Q3 (75th percentile)
- **4**: Maximum

**Examples:**
```excel
=QUARTILE.INC(A1:A100, 1)          → First quartile
=QUARTILE.INC(A1:A100, 3)          → Third quartile
```

**IQR (Interquartile Range):**
```excel
=QUARTILE.INC(A:A, 3) - QUARTILE.INC(A:A, 1)
```

**Use:** Detect outliers

---

### LARGE
**Returns kth largest value**

**Syntax:** `=LARGE(array, k)`

**Examples:**
```excel
=LARGE(A1:A100, 1)                 → Largest value
=LARGE(A1:A100, 2)                 → 2nd largest
=LARGE(A1:A100, 10)                → 10th largest
```

**Use Cases:**
- Top performers
- Highest sales
- Best scores

---

### SMALL
**Returns kth smallest value**

**Syntax:** `=SMALL(array, k)`

**Examples:**
```excel
=SMALL(A1:A100, 1)                 → Smallest value
=SMALL(A1:A100, 5)                 → 5th smallest
```

---

## Advanced Statistical Functions

### COUNT, COUNTA, COUNTBLANK

**COUNT** - Counts cells with numbers
```excel
=COUNT(A1:A10)                     → Count numeric cells
```

**COUNTA** - Counts non-empty cells
```excel
=COUNTA(A1:A10)                    → Count all non-empty
```

**COUNTBLANK** - Counts empty cells
```excel
=COUNTBLANK(A1:A10)                → Count blanks
```

---

### MAX & MIN
**Maximum and minimum values**

**Syntax:**
```excel
=MAX(number1, [number2], ...)
=MIN(number1, [number2], ...)
```

**Examples:**
```excel
=MAX(A1:A10)                       → Largest value
=MIN(A1:A10)                       → Smallest value
=MAX(A1:A10, B1:B10)               → Max across multiple ranges
```

---

### MAXIFS & MINIFS
**Conditional max/min (Excel 2019+)**

**Syntax:**
```excel
=MAXIFS(max_range, criteria_range1, criteria1, ...)
=MINIFS(min_range, criteria_range1, criteria1, ...)
```

**Examples:**
```excel
=MAXIFS(Sales, Region, "North", Product, "Apple")
=MINIFS(Price, Category, "Electronics", Stock, ">0")
```

---

### FREQUENCY
**Distribution frequency (array function)**

**Syntax:** `=FREQUENCY(data_array, bins_array)`

**Use:** Create histogram data

**Example:**
```excel
// Bins: 10, 20, 30, 40, 50
=FREQUENCY(A1:A100, B1:B5)
```

**Returns:** Count in each bin

---

### CONFIDENCE.NORM
**Confidence interval**

**Syntax:** `=CONFIDENCE.NORM(alpha, standard_dev, size)`

**Example:**
```excel
=CONFIDENCE.NORM(0.05, 2.5, 50)    → 95% confidence interval
```

---

### SKEW
**Skewness of distribution**

**Syntax:** `=SKEW(number1, [number2], ...)`

**Returns:**
- **Positive**: Right-skewed (tail on right)
- **Negative**: Left-skewed (tail on left)
- **Zero**: Symmetric

---

### KURT
**Kurtosis (peakedness)**

**Syntax:** `=KURT(number1, [number2], ...)`

**Use:** Measure distribution shape

---

## Practical Examples

### Calculate Grade Statistics
```excel
=AVERAGE(Scores)                   → Mean
=MEDIAN(Scores)                    → Median
=STDEV.S(Scores)                   → Standard deviation
=MIN(Scores)                       → Lowest score
=MAX(Scores)                       → Highest score
=PERCENTILE.INC(Scores, 0.75)      → 75th percentile
```

### Identify Outliers
```excel
// Values beyond 1.5 * IQR from quartiles
// Assuming data in A1:A100, testing value in B1

Q1 (in C1): =QUARTILE.INC($A$1:$A$100, 1)
Q3 (in D1): =QUARTILE.INC($A$1:$A$100, 3)
IQR (in E1): =D1-C1
Lower (in F1): =C1-1.5*E1
Upper (in G1): =D1+1.5*E1
Outlier Check (in H1): =OR(B1<F1, B1>G1)
```

### Quality Control (Six Sigma)
```excel
Mean: =AVERAGE(Measurements)
StdDev: =STDEV.S(Measurements)
UCL: =Mean + 3*StdDev             → Upper control limit
LCL: =Mean - 3*StdDev             → Lower control limit
```

### Sales Forecast
```excel
=FORECAST.LINEAR(NextMonth, Sales, Months)
=TREND(Sales, Months, NextMonths)
```

---

## Quick Reference

| Function | Purpose | Example |
|----------|---------|---------|
| AVERAGE | Mean | `=AVERAGE(A1:A10)` |
| MEDIAN | Middle value | `=MEDIAN(A1:A10)` |
| MODE.SNGL | Most common | `=MODE.SNGL(A1:A10)` |
| STDEV.S | Std deviation | `=STDEV.S(A1:A10)` |
| VAR.S | Variance | `=VAR.S(A1:A10)` |
| CORREL | Correlation | `=CORREL(A:A,B:B)` |
| RANK.EQ | Rank | `=RANK.EQ(A2,A:A,0)` |
| PERCENTILE | Percentile | `=PERCENTILE.INC(A:A,0.75)` |
| LARGE | Kth largest | `=LARGE(A:A,5)` |
| SMALL | Kth smallest | `=SMALL(A:A,5)` |
| FORECAST | Forecast | `=FORECAST.LINEAR(x,y,x)` |

---

**[⬆ Back to Main README](../../README.md)**
