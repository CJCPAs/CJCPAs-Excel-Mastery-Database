# Text Functions

> **Manipulate, extract, and format text data**

## Function Quick Reference

| Function | Purpose | Example |
|----------|---------|---------|
| **LEFT** | First N characters | `=LEFT("Hello",2)` → "He" |
| **RIGHT** | Last N characters | `=RIGHT("Hello",2)` → "lo" |
| **MID** | Extract from middle | `=MID("Hello",2,3)` → "ell" |
| **LEN** | Text length | `=LEN("Hello")` → 5 |
| **FIND** | Position (case-sensitive) | `=FIND("l","Hello")` → 3 |
| **SEARCH** | Position (case-insensitive) | `=SEARCH("L","Hello")` → 3 |
| **SUBSTITUTE** | Replace text | `=SUBSTITUTE("Hi","i","o")` → "Ho" |
| **REPLACE** | Replace by position | `=REPLACE("ABC",2,1,"X")` → "AXC" |
| **TRIM** | Remove extra spaces | `=TRIM("  Hi  ")` → "Hi" |
| **CLEAN** | Remove non-printable | `=CLEAN(A1)` |
| **UPPER** | Uppercase | `=UPPER("hi")` → "HI" |
| **LOWER** | Lowercase | `=LOWER("HI")` → "hi" |
| **PROPER** | Title case | `=PROPER("john doe")` → "John Doe" |
| **CONCAT** | Join text | `=CONCAT("A","B")` → "AB" |
| **TEXTJOIN** | Join with delimiter | `=TEXTJOIN(",",TRUE,A1:A3)` |
| **TEXT** | Format as text | `=TEXT(123,"0000")` → "0123" |
| **VALUE** | Text to number | `=VALUE("123")` → 123 |
| **REPT** | Repeat text | `=REPT("*",5)` → "*****" |
| **EXACT** | Case-sensitive compare | `=EXACT("Hi","hi")` → FALSE |
| **CHAR** | Character from code | `=CHAR(65)` → "A" |
| **CODE** | Code from character | `=CODE("A")` → 65 |
| **CONCATENATE** | Join (legacy) | `=CONCATENATE(A1,B1)` |

## Common Solutions

### Split Names
```excel
First: =LEFT(A1,FIND(" ",A1)-1)
Last:  =RIGHT(A1,LEN(A1)-FIND(" ",A1))
```

### Clean Phone Numbers
```excel
=SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(A1,"-",""),"(",""),")","")
```

### Extract Domain from Email
```excel
=MID(A1,FIND("@",A1)+1,100)
```

### Add Leading Zeros
```excel
=TEXT(A1,"00000")
```

---

[📚 Full Text Solutions](../../solutions/text-manipulation/) | [🏠 Back to Home](../../README.md)
