# Financial Functions

> **Loans, investments, depreciation, and financial analysis**

## Function Quick Reference

| Function | Purpose | Example |
|----------|---------|---------|
| **PMT** | Loan payment | `=PMT(6%/12,360,-250000)` |
| **PPMT** | Principal portion | `=PPMT(rate,period,nper,pv)` |
| **IPMT** | Interest portion | `=IPMT(rate,period,nper,pv)` |
| **PV** | Present value | `=PV(rate,nper,pmt)` |
| **FV** | Future value | `=FV(rate,nper,pmt,pv)` |
| **NPER** | Number of periods | `=NPER(rate,pmt,pv)` |
| **RATE** | Interest rate | `=RATE(nper,pmt,pv)` |
| **NPV** | Net present value | `=NPV(rate,values)` |
| **XNPV** | NPV with dates | `=XNPV(rate,values,dates)` |
| **IRR** | Internal rate of return | `=IRR(values)` |
| **XIRR** | IRR with dates | `=XIRR(values,dates)` |
| **MIRR** | Modified IRR | `=MIRR(values,fin_rate,reinv_rate)` |
| **SLN** | Straight-line depreciation | `=SLN(cost,salvage,life)` |
| **DB** | Declining balance | `=DB(cost,salvage,life,period)` |
| **DDB** | Double declining balance | `=DDB(cost,salvage,life,period)` |
| **SYD** | Sum-of-years digits | `=SYD(cost,salvage,life,period)` |
| **EFFECT** | Effective rate | `=EFFECT(nominal,periods)` |
| **NOMINAL** | Nominal rate | `=NOMINAL(effective,periods)` |
| **CUMIPMT** | Cumulative interest | `=CUMIPMT(rate,nper,pv,start,end,type)` |
| **CUMPRINC** | Cumulative principal | `=CUMPRINC(rate,nper,pv,start,end,type)` |

## Common Solutions

### Monthly Loan Payment
```excel
=PMT(AnnualRate/12, Years*12, -LoanAmount)
```

### How Much Can I Borrow?
```excel
=PV(Rate/12, Term*12, -MonthlyPayment)
```

### Investment Growth
```excel
=FV(Rate/12, Years*12, -MonthlyContribution, -InitialAmount)
```

### Payoff Time
```excel
=NPER(Rate/12, -Payment, Balance)
```

---

[📚 Full Financial Solutions](../../solutions/financial/) | [🏠 Back to Home](../../README.md)
