# Jumia Kenya Product Pricing, Discounts and Ratings Analysis (Excel)

An Excel data analysis project that looks at 109 discounted products listed on Jumia Kenya. It asks how deeply products are discounted, how customers rate them, and whether bigger discounts lead to better ratings or more customer engagement. The workbook contains the cleaned dataset, a set of pivot-table analyses, and an interactive dashboard.

## Business questions

1. What is the average current price, old price, discount and rating?
2. How are products spread across discount levels and rating levels?
3. Do higher discounts lead to more customer engagement (reviews)?
4. Do higher-rated products tend to have more reviews?
5. Which products have the biggest discounts, the most reviews and the best ratings?
6. How do high-discount and low-discount products compare on rating and reviews?

## Dataset

109 products, one row each, in an Excel table (`Excel_jumia` sheet).

| Column | Description |
|---|---|
| `Product` | Product name as listed on Jumia |
| `Current price` | Selling price (KSh) |
| `old price` | Original price before the discount (KSh) |
| `Discount` | Discount percentage |
| `Review` | Number of customer reviews, or `Unknown` if the listing had none |
| `Ratings` | Star rating (out of 5), or `Unknown` if the listing had none |

**Calculated columns added in Excel**

| Column | Formula idea | Purpose |
|---|---|---|
| `Positive Reviews` | `=IFERROR(ABS(Review), Review)` | Review counts came in as negative numbers (for example `-5`), so the absolute value gives the real count and leaves `Unknown` as it is |
| `Discount Amount (KSh)` | `=Old price - Current price` | How many shillings a customer saves |
| `Rating Category` | `Unknown`, `Poor` (below 3), `Average` (3 to 4.4), `Excellent` (above 4.4) | Groups ratings into bands |
| `Discount Category` | `Low` (under 20%), `Medium` (20% to 40%), `High` (above 40%) | Groups discounts into bands |

## Method

1. **Cleaned and structured the data:** converted the range to an Excel table, checked for duplicates (none) and handled the `Unknown` values and negative review counts.
2. **Engineered features** with Excel formulas (the four calculated columns above).
3. **Built pivot tables** for each business question, one per sheet.
4. **Built a dashboard** from six charts with an interactive slicer, so the dashboard can be filtered by category.

## Workbook contents

| Sheet | What it answers |
|---|---|
| `Excel_jumia` | Cleaned data and calculated columns |
| `Product Discount Categories` | Number of products per discount category |
| `Averages` | Average price, old price, discount and rating |
| `Average Rating of Discount` | Average discount and rating by discount category |
| `Product Distribution Rating Ca...` | Number of products per rating category |
| `Discount % vs Reviews`, `Discount % C.Engagement` | Reviews at each discount level |
| `Rating vs Reviews`, `Higher Rating vs Reviews` | Reviews at each rating level or category |
| `Top 10 Discounts`, `Top 10 Reviews` | Biggest discounts and most-reviewed products |
| `Top 5 Highest rated products`, `Bottom 5 lowest-rated products` | Best and worst rated products |
| `Discount, Reviews, Ratings` | High-discount against low-discount comparison |
| `Dashboard` | Six charts and a slicer in one view |

## Key findings

**Pricing and discounts**
- The average product sells for **KSh 1,208**, down from an average old price of **KSh 1,838**. The average discount is **36.4%**, and the typical product is discounted by KSh 629.
- Prices range from KSh 38 to KSh 3,750. The most expensive item is a 32-piece cordless drill set and the cheapest is a 3-piece knitting needle set.
- Heavy discounting is the norm: **59 of 109 products (54%) have a high discount** (over 40%), 32 have a medium discount and only 18 have a low discount.
- The deepest discounts are 64% (a 6-in-1 bottle opener) and 61% (an owl keychain).

**Ratings**
- Nearly half of the products (**52 of 109**) have no rating or review yet, so every rating-based finding below rests on the 57 products that do.
- Among those 57, the average rating is **3.89 out of 5**: 23 rated Excellent, 22 Average and 12 Poor. Seven products have a perfect 5.0.

**Do bigger discounts mean better ratings or more engagement?**
- **Medium-discount products are rated best**, at an average of 4.28, compared with 3.61 for high-discount and 3.73 for low-discount products. The low-discount figure rests on only 4 rated products, so it should be treated with caution.
- Medium-discount products also draw the most reviews per product (about 15, against about 11 for high discount). Across all products a deeper discount shows a **weak negative** link with review count (correlation about -0.14), not a positive one.
- **Rating and review count are essentially unrelated** (correlation about 0.06). The most-reviewed product, a 120W cordless vacuum cleaner with 69 reviews, has a below-average rating of 2.8.

**Takeaway:** in this sample, steeper discounts did not bring better ratings or more reviews. A moderate discount of 20% to 40% performed best on both.

## Limitations

- **Small sample:** 109 products, and only 57 with ratings and reviews. Correlations on 57 points are indicative, not conclusive.
- **A snapshot:** the data is a single listing capture and has no sales figures. Reviews are a stand-in for engagement, not for purchases.
- **Counts versus totals:** the discount-versus-reviews pivots count the number of products at each discount level and do not add up the reviews. A sum or average of reviews per discount level would answer the engagement question more directly.
- **Lowest-rated products:** the "Bottom 5" pivot shows ratings of 0 because products with an `Unknown` rating are treated as zero by the Max aggregation. Those products are unrated, not zero-rated. The real lowest rating in the data is 2.0.

## Tools

Microsoft Excel: tables, formulas (`IF`, `IFERROR`, `ABS`, `SUBTOTAL`), pivot tables, pivot charts, slicers and dashboard design.

## How to use

1. Download `excell_jumia.xlsx` from this repository.
2. Open it in Excel and start on the `Dashboard` sheet.
3. Use the slicer to filter, or open any analysis sheet to see its pivot table.

## Author

**Ephantus Macharia**, aspiring data analyst, Nairobi, Kenya
[GitHub](https://github.com/machariaephantus96-coder)
