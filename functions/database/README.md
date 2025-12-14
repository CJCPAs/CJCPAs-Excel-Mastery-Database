# Database Functions

> **Perform calculations on database-style tables with criteria**

## Function Quick Reference

| Function | Purpose | Syntax |
|----------|---------|--------|
| **DSUM** | Sum with criteria | `=DSUM(database, field, criteria)` |
| **DAVERAGE** | Average with criteria | `=DAVERAGE(database, field, criteria)` |
| **DCOUNT** | Count numbers | `=DCOUNT(database, field, criteria)` |
| **DCOUNTA** | Count non-empty | `=DCOUNTA(database, field, criteria)` |
| **DMAX** | Maximum | `=DMAX(database, field, criteria)` |
| **DMIN** | Minimum | `=DMIN(database, field, criteria)` |
| **DGET** | Single value | `=DGET(database, field, criteria)` |
| **DPRODUCT** | Product | `=DPRODUCT(database, field, criteria)` |
| **DSTDEV** | Sample std dev | `=DSTDEV(database, field, criteria)` |
| **DSTDEVP** | Population std dev | `=DSTDEVP(database, field, criteria)` |
| **DVAR** | Sample variance | `=DVAR(database, field, criteria)` |
| **DVARP** | Population variance | `=DVARP(database, field, criteria)` |

## How Database Functions Work

**Database:** Range including headers (A1:D100)
**Field:** Column name ("Sales") or number (2)
**Criteria:** Range with header + conditions

### Criteria Range Example
| Region | Sales |
|--------|-------|
| North | >1000 |

## Example
```excel
=DSUM(A1:D100, "Sales", F1:G2)
```
Sums Sales where criteria match.

### Multiple Criteria (AND)
| Region | Sales |
|--------|-------|
| North | >1000 |

### Multiple Criteria (OR)
| Region |
|--------|
| North |
| South |

---

[🏠 Back to Home](../../README.md)
