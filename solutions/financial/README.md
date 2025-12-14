# Financial Calculations & Analysis Solutions

> **Loans, investments, depreciation, and financial modeling**

## Quick Solutions

| I want to... | Solution |
|--------------|----------|
| Calculate loan payment | [PMT Function](#loan-payment-pmt) |
| Figure out how much I can borrow | [Loan Amount from Payment](#calculate-loan-amount) |
| Calculate investment growth | [Future Value](#future-value-of-investment) |
| Find present value | [Present Value](#present-value-pv) |
| Calculate ROI/IRR | [Internal Rate of Return](#internal-rate-of-return-irr) |
| Compare investments with NPV | [Net Present Value](#net-present-value-npv) |
| Calculate depreciation | [Depreciation Methods](#depreciation) |
| Build an amortization schedule | [Loan Amortization](#amortization-schedule) |
| Calculate compound interest | [Compound Interest](#compound-interest) |
| Find break-even point | [Break-Even Analysis](#break-even-analysis) |

---

## Loan Payment (PMT)

### The Challenge
Calculate the monthly payment for a loan given principal, interest rate, and term.

### Quick Answer
```excel
=PMT(rate/12, term*12, -principal)
```

### Full Example

**Loan Details:**
| Item | Value |
|------|-------|
| Loan Amount | $250,000 |
| Annual Interest Rate | 6.5% |
| Term | 30 years |

**Formula:** `=PMT(6.5%/12, 30*12, -250000)`

**Result:** `$1,580.17` per month

**Explanation:**
- `6.5%/12` = monthly interest rate (0.5417%)
- `30*12` = total payments (360 months)
- `-250000` = negative because it's money you receive (then pay back)

### With Extra Payment Calculation
```excel
Monthly Payment: =PMT(B2/12, B3*12, -B1)
Total Paid:      =PMT(B2/12, B3*12, -B1) * B3 * 12
Total Interest:  =(PMT(B2/12, B3*12, -B1) * B3 * 12) + B1
```

---

## Calculate Loan Amount

### The Challenge
Figure out how much you can borrow based on what you can afford to pay monthly.

### Quick Answer
```excel
=PV(rate/12, term*12, -payment)
```

### Full Example

**What You Know:**
| Item | Value |
|------|-------|
| Can Afford Monthly | $2,000 |
| Interest Rate | 5.5% |
| Term | 15 years |

**Formula:** `=PV(5.5%/12, 15*12, -2000)`

**Result:** `$245,578.58` (maximum loan amount)

---

## Future Value of Investment

### The Challenge
Calculate how much an investment will be worth after a period of time with regular contributions.

### Quick Answer
```excel
=FV(rate/periods, total_periods, -payment, -initial_investment, type)
```

### Full Example - Retirement Savings

**Investment Plan:**
| Item | Value |
|------|-------|
| Initial Investment | $10,000 |
| Monthly Contribution | $500 |
| Annual Return | 7% |
| Years | 25 |

**Formula:** `=FV(7%/12, 25*12, -500, -10000, 0)`

**Result:** `$459,803.78`

**Breakdown:**
- Initial $10,000 grows to ~$54,000
- Monthly $500 contributions grow to ~$406,000
- Total: ~$460,000

### Simple Future Value (No Contributions)
```excel
=FV(rate, periods, 0, -principal)
```

**Example:** $10,000 at 5% for 10 years
```excel
=FV(5%, 10, 0, -10000)  → $16,288.95
```

---

## Present Value (PV)

### The Challenge
Calculate how much a future amount is worth today (or how much to invest now).

### Quick Answer
```excel
=PV(rate, periods, 0, -future_value)
```

### Full Example

**Goal:** Have $100,000 in 10 years at 6% annual return

**Formula:** `=PV(6%, 10, 0, -100000)`

**Result:** `$55,839.48` (invest this today)

### Present Value of Annuity
How much is a series of future payments worth today?

**Example:** Receive $1,000/month for 5 years, 5% annual rate
```excel
=PV(5%/12, 5*12, -1000)  → $52,990.71
```

---

## Internal Rate of Return (IRR)

### The Challenge
Find the actual return rate of an investment with irregular cash flows.

### Quick Answer
```excel
=IRR(cash_flows)
```

### Full Example

**Investment Cash Flows:**
| Year | Cash Flow | Description |
|------|-----------|-------------|
| 0 | -$50,000 | Initial investment |
| 1 | $15,000 | Year 1 return |
| 2 | $18,000 | Year 2 return |
| 3 | $20,000 | Year 3 return |
| 4 | $22,000 | Year 4 return |

**Data in A1:A5:** `-50000, 15000, 18000, 20000, 22000`

**Formula:** `=IRR(A1:A5)`

**Result:** `19.44%` annual return

### XIRR for Specific Dates
When cash flows occur on specific dates:

| Date | Cash Flow |
|------|-----------|
| 1/1/2024 | -$50,000 |
| 6/15/2024 | $10,000 |
| 12/1/2024 | $15,000 |
| 7/1/2025 | $35,000 |

```excel
=XIRR(B1:B4, A1:A4)
```

---

## Net Present Value (NPV)

### The Challenge
Determine if an investment is worth making by calculating its value in today's dollars.

### Quick Answer
```excel
=NPV(rate, future_cash_flows) + initial_investment
```

### Full Example

**Project Analysis:**
| Year | Cash Flow |
|------|-----------|
| 0 | -$100,000 (investment) |
| 1 | $30,000 |
| 2 | $35,000 |
| 3 | $40,000 |
| 4 | $45,000 |

**Discount Rate:** 8%

**Formula:** `=NPV(8%, B2:B5) + B1`

Where B1 = -100000, B2:B5 = future cash flows

**Result:** `$21,410.07`

**Decision:**
- NPV > 0 → Good investment
- NPV < 0 → Reject investment
- NPV = 0 → Break-even

### Compare Two Projects
| | Project A | Project B |
|--|-----------|-----------|
| Initial | -$50,000 | -$80,000 |
| Year 1 | $20,000 | $30,000 |
| Year 2 | $25,000 | $35,000 |
| Year 3 | $30,000 | $40,000 |
| **NPV (10%)** | **$11,037** | **$8,224** |

Project A has higher NPV → Better investment

---

## Depreciation

### Straight-Line Depreciation (SLN)
Equal depreciation each year.

```excel
=SLN(cost, salvage, life)
```

**Example:** $50,000 asset, $5,000 salvage, 10-year life
```excel
=SLN(50000, 5000, 10)  → $4,500 per year
```

### Declining Balance (DB)
Accelerated depreciation - more in early years.

```excel
=DB(cost, salvage, life, period)
```

**Example:** Year 1 depreciation
```excel
=DB(50000, 5000, 10, 1)  → $9,500
```

### Double Declining Balance (DDB)
Fastest depreciation method.

```excel
=DDB(cost, salvage, life, period)
```

**Example:** Year 1
```excel
=DDB(50000, 5000, 10, 1)  → $10,000
```

### Sum-of-Years Digits (SYD)
```excel
=SYD(cost, salvage, life, period)
```

### Depreciation Schedule Example

| Year | SLN | DB | DDB | SYD |
|------|-----|-----|-----|-----|
| 1 | $4,500 | $9,500 | $10,000 | $8,182 |
| 2 | $4,500 | $7,695 | $8,000 | $7,364 |
| 3 | $4,500 | $6,233 | $6,400 | $6,545 |
| ... | ... | ... | ... | ... |

---

## Amortization Schedule

### Build Complete Schedule

**Setup:**
| A | B |
|---|---|
| Loan Amount | $200,000 |
| Annual Rate | 6% |
| Years | 30 |
| Monthly Payment | =PMT(B2/12,B3*12,-B1) |

**Schedule (starting row 7):**
| Payment # | Payment | Principal | Interest | Balance |
|-----------|---------|-----------|----------|---------|
| 1 | $1,199.10 | $199.10 | $1,000.00 | $199,800.90 |

**Formulas:**
```excel
Payment #:    =ROW()-6  (or 1, 2, 3...)
Payment:      =$B$4  (fixed monthly payment)
Interest:     =E6*$B$2/12  (previous balance × monthly rate)
Principal:    =B7-C7  (payment - interest)
Balance:      =E6-D7  (previous balance - principal)
```

**Copy formulas down for 360 rows**

### Quick Payoff Calculator
Months to pay off with extra payment:
```excel
=NPER(rate/12, -(payment + extra), balance)
```

---

## Compound Interest

### Basic Compound Interest
```excel
=principal * (1 + rate/n)^(n*years)
```

Where n = compounding periods per year

### Example
$10,000 at 5% for 10 years, compounded:

| Frequency | n | Formula | Result |
|-----------|---|---------|--------|
| Annually | 1 | `=10000*(1+5%/1)^(1*10)` | $16,288.95 |
| Quarterly | 4 | `=10000*(1+5%/4)^(4*10)` | $16,436.19 |
| Monthly | 12 | `=10000*(1+5%/12)^(12*10)` | $16,470.09 |
| Daily | 365 | `=10000*(1+5%/365)^(365*10)` | $16,486.65 |

### Continuous Compounding
```excel
=principal * EXP(rate * years)
```

---

## Break-Even Analysis

### The Challenge
Find the sales volume needed to cover all costs.

### Quick Answer
```excel
=Fixed_Costs / (Price - Variable_Cost)
```

### Full Example

| Item | Value |
|------|-------|
| Fixed Costs | $50,000 |
| Price per Unit | $25 |
| Variable Cost per Unit | $10 |

**Formula:** `=50000/(25-10)`

**Result:** `3,333.33` units to break even

### Break-Even Revenue
```excel
=Fixed_Costs / (1 - Variable_Cost/Price)
=50000 / (1 - 10/25)  → $83,333.33
```

### Break-Even with Target Profit
```excel
=(Fixed_Costs + Target_Profit) / (Price - Variable_Cost)
=(50000 + 20000) / (25 - 10)  → 4,667 units
```

---

## Financial Ratios

### Profitability Ratios
```excel
Gross Margin:       =(Revenue - COGS) / Revenue
Net Margin:         =Net_Income / Revenue
ROE:                =Net_Income / Shareholders_Equity
ROA:                =Net_Income / Total_Assets
```

### Liquidity Ratios
```excel
Current Ratio:      =Current_Assets / Current_Liabilities
Quick Ratio:        =(Current_Assets - Inventory) / Current_Liabilities
```

### Leverage Ratios
```excel
Debt-to-Equity:     =Total_Debt / Total_Equity
Interest Coverage:  =EBIT / Interest_Expense
```

---

## Interest Rate Conversions

### Nominal to Effective (Annual)
```excel
=EFFECT(nominal_rate, periods_per_year)
```

**Example:** 6% compounded monthly
```excel
=EFFECT(6%, 12)  → 6.17% effective annual rate
```

### Effective to Nominal
```excel
=NOMINAL(effective_rate, periods_per_year)
```

---

## Common Financial Errors & Fixes

| Error | Cause | Fix |
|-------|-------|-----|
| #NUM! in IRR | No sign change in cash flows | Ensure negative initial investment |
| Wrong PMT result | Rate/period mismatch | Use rate/12 for monthly payments |
| Negative result | Sign confusion | Check if principal should be negative |
| #VALUE! | Text in numbers | Ensure all values are numeric |

---

## Pro Tips

1. **Consistent Time Periods:** Match rate and nper (annual rate with annual periods, or monthly rate with monthly periods)

2. **Sign Conventions:**
   - Money you pay out = negative
   - Money you receive = positive

3. **Type Argument:**
   - 0 = payment at end of period (default)
   - 1 = payment at beginning of period

4. **Validation:** Cross-check with online calculators for complex scenarios

---

## Related Solutions

- [Date & Time Calculations](../dates-times/README.md) - For payment dates
- [Conditional Calculations](../conditional-calculations/README.md) - For financial criteria
- [Data Analysis](../data-analysis/README.md) - For financial modeling

---

[🏠 Back to Home](../../README.md) | [🎯 All Solutions](../README.md)
