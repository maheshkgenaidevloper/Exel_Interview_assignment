::: {align="center"}
# 📊 EXEL --- MS Excel Assignment

**Quick Answers • Data Analysis • Excel Essentials**

*A concise guide to commonly asked Microsoft Excel questions.*

![Excel](https://img.shields.io/badge/Microsoft-Excel-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)
![Topic](https://img.shields.io/badge/Focus-Data%20Analysis-2563EB?style=for-the-badge)
![Format](https://img.shields.io/badge/Answers-Short%20%26%20Simple-7C3AED?style=for-the-badge)
:::

------------------------------------------------------------------------

## 🧩 Excel Fundamentals

### 1. How many data formats are available in Excel? Name some.

Excel supports several cell number formats, including **General, Number,
Currency, Accounting, Date, Time, Percentage, Fraction, Scientific, and
Text**.

**Example:** Format `0.25` as Percentage to display `25%`.

### 2. What are charts in MS Excel?

Charts are visual representations of data that make comparisons,
patterns, and trends easier to understand.

**Example:** Use a column chart to compare monthly sales.

### 3. Explain Pivot Tables and their uses.

A PivotTable summarizes and analyzes large datasets by grouping,
counting, or calculating values without changing the original data.

**Example:** Summarize total sales by region.

### 4. How would you provide a dynamic range in the Data Source of PivotTables?

Convert the source data into an **Excel Table** using **Ctrl + T**, then
build the PivotTable from that table. New rows are included when you
refresh it.

**Example:** Use `SalesTable` as the PivotTable source and select
**Refresh** after adding sales rows.

### 5. Is it possible to make a PivotTable using multiple data sources?

Yes. You can use the **Data Model/Power Pivot** to relate multiple
tables, or combine sources with Power Query first.

**Example:** Relate `Sales` and `Products` tables using `ProductID`.

### 6. Difference between COUNT, COUNTA, COUNTIF, and COUNTBLANK

-   **COUNT:** Counts cells containing numbers.
-   **COUNTA:** Counts non-empty cells.
-   **COUNTIF:** Counts cells that meet a condition.
-   **COUNTBLANK:** Counts empty cells.

**Example:** `=COUNTIF(A2:A10,">50")` counts values greater than 50.

### 7. What is the IF function in Excel?

`IF` checks a condition and returns one result if it is true and another
if it is false.

**Example:** `=IF(B2>=40,"Pass","Fail")`

### 8. What is the use of the LOOKUP function?

`LOOKUP` searches for a value in a row or column and returns a
corresponding value from another row or column. It generally expects
sorted lookup data.

**Example:** `=LOOKUP(75,A2:A5,B2:B5)` looks up 75 in column A and
returns the corresponding value from column B.

### 9. What is the "What-If" condition in Excel formulas?

What-If Analysis explores how changing input values affects a formula's
result. Excel tools include **Goal Seek, Scenario Manager, and Data
Tables**.

**Example:** Use Goal Seek to find the exam score needed to reach an
overall average of 80.

### 10. What is the default value of the last parameter of VLOOKUP?

The `range_lookup` argument defaults to **TRUE** when omitted, meaning
an approximate match is used. For an exact match, specify `FALSE`.

**Example:** `=VLOOKUP(E2,A2:C10,3,FALSE)`

------------------------------------------------------------------------

## 📈 Statistics & Visualization

### 11. How do you find duplicates in a column?

Use **Conditional Formatting → Highlight Cells Rules → Duplicate
Values**.

**Example:** Select `A2:A100` and highlight duplicate entries.

### 12. Difference between MEAN, MEDIAN, and MODE

-   **Mean:** The arithmetic average.
-   **Median:** The middle value after sorting.
-   **Mode:** The most frequently occurring value.

**Example:** For `2, 3, 3, 10`, mean = `4.5`, median = `3`, and mode =
`3`.

### 13. For which data types are MEAN, MEDIAN, and MODE most suitable?

-   **Mean:** Numerical interval/ratio data, especially without extreme
    outliers.
-   **Median:** Ordered or numerical data, especially when skewed or
    containing outliers.
-   **Mode:** Categorical or numerical data; useful for finding the most
    common category/value.

**Example:** Use the median for household incomes when a few incomes are
extremely high.

### 14. How do you detect outliers?

A common method is the **IQR rule**: values below `Q1 − 1.5 × IQR` or
above `Q3 + 1.5 × IQR` are potential outliers, where `IQR = Q3 − Q1`.

**Example:** If `Q1 = 10` and `Q3 = 20`, values below `−5` or above `35`
are potential outliers.

### 15. What is a Box Plot?

A Box Plot displays a dataset's median, quartiles, spread, and potential
outliers.

**Example:** Compare the distribution of test scores across two classes
with a Box & Whisker chart.

### 16. What is a Pareto Chart?

A Pareto Chart combines bars ordered from highest to lowest with a
cumulative-percentage line, helping identify the most significant
contributors.

**Example:** Rank customer complaints by type to see which types account
for most complaints.

### 17. What are the different methods to find correlation?

Correlation can be assessed using **Pearson's correlation** for linear
relationships between numerical variables and **Spearman's rank
correlation** for ranked or monotonic relationships. A scatter plot can
also help visualize the relationship.

**Example:** Use `=CORREL(A2:A20,B2:B20)` to calculate Pearson
correlation.

------------------------------------------------------------------------

## 🛠️ Excel Tools & Functions

### 18. What is Scenario Manager?

Scenario Manager stores and compares different sets of input values to
show how they affect a result.

**Example:** Compare profit under optimistic, expected, and pessimistic
sales scenarios.

### 19. What is the purpose of Goal Seek?

Goal Seek finds the input value needed to make a formula return a
specified result.

**Example:** Find the sales quantity needed to achieve a profit of
₹50,000.

### 20. Do you know any keyboard shortcuts?

Useful Excel shortcuts include: - **Ctrl + C / Ctrl + V:** Copy /
Paste - **Ctrl + Z:** Undo - **Ctrl + T:** Create a table - **Ctrl +
Shift + L:** Toggle filters - **Alt + =:** Insert AutoSum

**Example:** Press `Ctrl + T` to convert a selected data range into a
table.

### 21. What is RANDBETWEEN used for?

`RANDBETWEEN(bottom,top)` returns a random integer between the specified
lower and upper limits, inclusive. The result can change when Excel
recalculates.

**Example:** `=RANDBETWEEN(1,100)` generates a random whole number from
1 to 100.

### 22. What is the CHOOSE function in Excel?

`CHOOSE` returns a value from a list based on a numeric index.

**Example:** `=CHOOSE(2,"Red","Blue","Green")` returns `Blue`.

### 23. Does Power Query help in merging and appending datasets?

Yes. **Merge** joins tables using matching columns; **Append** stacks
rows from tables with compatible columns.

**Example:** Append January and February sales tables into one dataset.

### 24. Can Power Query transform data?

Yes. Power Query can clean, filter, split, rename, change data types,
remove duplicates, group, pivot, and unpivot data.

**Example:** Split a `Full Name` column into `First Name` and
`Last Name`.

------------------------------------------------------------------------

## ✅ Quick Notes

-   Refresh PivotTables after updating their source data.
-   Use `FALSE` in `VLOOKUP` when you need an exact match.
-   Check data types and missing values before analyzing data.
-   Correlation indicates association; it does not by itself prove
    causation.

------------------------------------------------------------------------

::: {align="center"}
**Made for the EXEL Assignment**\
*Learn it • Apply it • Analyze it*
:::
