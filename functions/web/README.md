# Web Functions

> **Retrieve and work with web data**

## Function Quick Reference

| Function | Purpose | Example |
|----------|---------|---------|
| **WEBSERVICE** | Get web data | `=WEBSERVICE(url)` |
| **FILTERXML** | Parse XML | `=FILTERXML(xml, xpath)` |
| **ENCODEURL** | URL encode | `=ENCODEURL(text)` |

## WEBSERVICE
Retrieves data from a web service.

```excel
=WEBSERVICE("https://api.example.com/data")
```

**Limitations:**
- Must be GET request
- Returns text up to 32,767 characters
- URL must be HTTPS
- May be blocked by firewalls

## FILTERXML
Extracts values from XML using XPath.

```excel
=FILTERXML(WEBSERVICE(url), "//element")
```

**Example - Get Title from XML:**
```excel
=FILTERXML(A1, "//title")
```

## ENCODEURL
Converts text to URL-safe format.

```excel
=ENCODEURL("hello world")
→ "hello%20world"
```

**Use case - Build API URL:**
```excel
="https://api.example.com/search?q=" & ENCODEURL(A1)
```

## Practical Example

### Get Stock Quote (Concept)
```excel
URL:      ="https://api.stockdata.com/quote/" & A1
Response: =WEBSERVICE(B1)
Parse:    =FILTERXML(C1, "//price")
```

## Alternative: Power Query
For complex web data, use:
Data → Get Data → From Web

---

[🏠 Back to Home](../../README.md)
