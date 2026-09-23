# Category Profitability (Excel)

Compare total profit against product category using the same retail sales dataset as the earlier "Basic Sales Summary" project. Grouped by category, with totals, averages, and a chart, built entirely with Excel formulas.

## Objective

Group profit by category, compare **totals and averages**, and identify the most profitable category. Deliverables: a **profit summary table** and a **chart**.

## Results

| Category | Total Profit | Average Profit | Transactions |
|---|---|---|---|
| Furniture | $193,585.57 | $202.28 | 957 |
| Office Supplies | $17,177.81 | $16.68 | 1,030 |
| Technology | $330,096.19 | $325.86 | 1,013 |
| **All categories** | **$540,859.57** | **$180.29** | **3,000** |

**Most profitable category: Technology** ($330,096.19, 61.0% of total profit)

![Total Profit by Category](chart.png)

### Key findings

- **Technology leads on total profit**, at 61.0% of all profit, but its **margin** (profit ÷ sales, 17.5%) is only slightly ahead of Furniture and Office Supplies (~16.8% each). Its lead comes mainly from higher sales volume, not a much better margin.
- **Office Supplies has the most transactions** (1,030) of any category but the lowest profit by far — an average of just $16.68 per sale, about 20× less than Technology.
- **423 of 3,000 transactions (14.1%) show a negative profit**, spread fairly evenly across all three categories.
- Within Technology, **Copiers** alone contribute $223,811.64 in profit — most of the category's result.

## Files

| File | What it is |
|---|---|
| `Sales_Data.xlsx` | Workbook. New sheet **Category Profitability**, plus **Basic Sales Summary**, **Product Count Analysis**, and the original **KPI Summary** / **Sales Data** sheets (all unchanged). |


## The summary sheet

**Category Profitability** has:

1. **Profit summary table** (A5:D9): Category, Total Profit, Average Profit, Transactions, plus an "All categories" row.
2. **Most Profitable Category** card, found with INDEX/MATCH.
3. **Chart**: a column chart of Total Profit by Category.
4. A **VERIFICATION** panel underneath the table.

### Formulas used

Per category (row 6 shown; T = Profit, N = Category):

```excel
Total Profit    =SUMIFS('Sales Data'!$T$2:$T$5000,'Sales Data'!$N$2:$N$5000,$A6)
Average Profit  =AVERAGEIFS('Sales Data'!$T$2:$T$5000,'Sales Data'!$N$2:$N$5000,$A6)
Transactions    =COUNTIFS('Sales Data'!$N$2:$N$5000,$A6)
```

Overall:

```excel
Total Profit    =SUM('Sales Data'!T2:T5000)
Average Profit  =AVERAGE('Sales Data'!T2:T5000)
```

Most profitable category:

```excel
=INDEX(A6:A8,MATCH(MAX(B6:B8),B6:B8,0))
```


## Skills practised

`SUMIFS` · `AVERAGEIFS` · `COUNTIFS` · `INDEX`/`MATCH` · grouping and comparing categories · building a chart from a summary table · cross-checking totals

## Tools

Microsoft Excel
