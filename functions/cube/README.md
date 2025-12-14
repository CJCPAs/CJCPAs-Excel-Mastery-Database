# Cube Functions

> **Work with OLAP cubes and data models**

## Function Quick Reference

| Function | Purpose |
|----------|---------|
| **CUBEVALUE** | Return value from cube |
| **CUBEMEMBER** | Return member from cube |
| **CUBESET** | Define calculated set |
| **CUBERANKEDMEMBER** | Return nth member |
| **CUBEMEMBERPROPERTY** | Return member property |
| **CUBESETCOUNT** | Count members in set |
| **CUBEKPIMEMBER** | Return KPI property |

## Syntax Examples

### CUBEVALUE
```excel
=CUBEVALUE("connection", "member1", "member2", ...)
=CUBEVALUE("SalesCube", "[Measures].[Revenue]", "[Time].[2025]")
```

### CUBEMEMBER
```excel
=CUBEMEMBER("connection", "member_expression")
=CUBEMEMBER("SalesCube", "[Product].[Category].&[Electronics]")
```

### CUBESET
```excel
=CUBESET("connection", "set_expression", "caption")
=CUBESET("SalesCube", "[Product].[Category].Members", "All Categories")
```

### CUBERANKEDMEMBER
```excel
=CUBERANKEDMEMBER("connection", set, rank)
=CUBERANKEDMEMBER("SalesCube", TopProducts, 1)
```

## Requirements
- Connection to OLAP data source
- Analysis Services, PowerPivot, or compatible cube
- Knowledge of MDX expressions

## Use Cases
- Dynamic PivotTable alternatives
- Custom cube queries
- KPI dashboards from OLAP

---

[🏠 Back to Home](../../README.md)
