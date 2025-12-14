# Statistical Functions

> **Analyze data with averages, counts, rankings, and distributions**

## Function Quick Reference

| Function | Purpose | Example |
|----------|---------|---------|
| **AVERAGE** | Arithmetic mean | `=AVERAGE(A1:A100)` |
| **AVERAGEIF** | Conditional average | `=AVERAGEIF(A:A,">0",B:B)` |
| **AVERAGEIFS** | Multi-criteria average | `=AVERAGEIFS(C:C,A:A,"North",B:B,">100")` |
| **MEDIAN** | Middle value | `=MEDIAN(A1:A100)` |
| **MODE.SNGL** | Most frequent | `=MODE.SNGL(A1:A100)` |
| **COUNT** | Count numbers | `=COUNT(A1:A100)` |
| **COUNTA** | Count non-empty | `=COUNTA(A1:A100)` |
| **COUNTBLANK** | Count empty | `=COUNTBLANK(A1:A100)` |
| **COUNTIF** | Conditional count | `=COUNTIF(A:A,">100")` |
| **COUNTIFS** | Multi-criteria count | `=COUNTIFS(A:A,"North",B:B,">50")` |
| **MAX** | Maximum value | `=MAX(A1:A100)` |
| **MAXIFS** | Conditional max | `=MAXIFS(B:B,A:A,"North")` |
| **MIN** | Minimum value | `=MIN(A1:A100)` |
| **MINIFS** | Conditional min | `=MINIFS(B:B,A:A,"North")` |
| **LARGE** | Nth largest | `=LARGE(A1:A100,3)` |
| **SMALL** | Nth smallest | `=SMALL(A1:A100,3)` |
| **RANK.EQ** | Rank value | `=RANK.EQ(A1,$A:$A)` |
| **PERCENTILE.INC** | Percentile value | `=PERCENTILE.INC(A:A,0.9)` |
| **PERCENTRANK.INC** | Rank as percentile | `=PERCENTRANK.INC(A:A,A1)` |
| **QUARTILE.INC** | Quartile value | `=QUARTILE.INC(A:A,1)` |
| **STDEV.S** | Standard deviation | `=STDEV.S(A1:A100)` |
| **VAR.S** | Variance | `=VAR.S(A1:A100)` |
| **CORREL** | Correlation | `=CORREL(A:A,B:B)` |

## Common Solutions

### Five-Number Summary
```excel
Min:    =MIN(A:A)
Q1:     =QUARTILE.INC(A:A,1)
Median: =MEDIAN(A:A)
Q3:     =QUARTILE.INC(A:A,3)
Max:    =MAX(A:A)
```

### Count Unique Values
```excel
=SUMPRODUCT(1/COUNTIF(A2:A100,A2:A100))
```

### Outlier Detection (Z-Score)
```excel
=(A1-AVERAGE($A:$A))/STDEV.S($A:$A)
```

---

[📚 Full Data Analysis Solutions](../../solutions/data-analysis/) | [🏠 Back to Home](../../README.md)
