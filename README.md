# product-count-analysis
Product Count Analysis using Microsoft Excel and the Superstore dataset. This project analyzes product categories using COUNTIF and identifies the top category. It also calculates unique product counts for Furniture, Office Supplies, and Technology.
# Task 15: Product Count Analysis

## Project Overview
This project analyzes product categories using Microsoft Excel and the Superstore dataset. The objective is to count category records, identify the category with the highest count, and calculate the number of unique product names in each category.

## Objectives
- Count records by product category.
- Practice Excel COUNTIF functions.
- Identify the category with the highest number of records.
- Calculate unique product counts.
- Present the results in formatted tables.

## Tools Used
- Microsoft Excel
- Superstore dataset
- Excel formulas: COUNTIF, INDEX, MATCH, MAX, FILTER, UNIQUE, and COUNTA

## Dataset
The project uses a Superstore dataset containing order information, product categories, product names, sales, quantities, discounts, and profits.

## Analysis Performed
1. Counted records in each product category.
2. Identified the category with the highest record count.
3. Calculated unique product names for each category.
4. Created formatted summary tables.
5. Added a conclusion based on the results.

## Results

### Product Count by Category

| Category | Product Count |
|---|---:|
| Furniture | 9,876 |
| Office Supplies | 31,273 |
| Technology | 10,141 |

### Unique Products by Category

| Category | Unique Products |
|---|---:|
| Furniture | 842 |
| Office Supplies | 2,071 |
| Technology | 876 |

## Key Finding
Office Supplies has the highest number of category records, with 31,273 records. It also has the highest number of unique product names among the three categories, with 2,071.

## Excel Formulas

### Count Records by Category
```excel
=COUNTIF(Sheet1!L:L,A2)
```

Replace `Sheet1` with the actual name of the worksheet containing the source data.

### Count Unique Product Names
```excel
=COUNTA(UNIQUE(FILTER(SuperStoreOrders!$N$2:$N$100000,(SuperStoreOrders!$L$2:$L$100000=A11)*(SuperStoreOrders!$N$2:$N$100000<>""))))
```

This formula counts distinct, nonblank product names for the category specified in cell A11.

## Conclusion
This project demonstrates how Excel formulas can be used to analyze product categories, count records, identify the leading category, and calculate unique product names. The results show that Office Supplies leads in both category record count and unique product names.

## Deliverables
- Excel workbook containing the analysis
- Product count table
- Unique product count table
- Top category identification
- Conclusion and results

## Author
Data Analytics Track – Task 15
