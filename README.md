# -DATA-EXPLORATION

Metric	Business Insight	Excel Formula	Values

Total Products	Total volume of tracked SKUs	=COUNTA(B2:B35)	34
Total Inventory Value	Cumulative sum of single unit prices	=SUM(D2:D35)	$10,100
Average Unit Price	Mean product price across catalog	=AVERAGE(D2:D35)	$297.06
Minimum Price	Lowest entry-level item price	=MIN(D2:D35)	$30.00
Maximum Price	Highest premium item price	=MAX(D2:D35)	$1,000.00
Electronics Category Value	Total price value of Electronics items	=SUMIF(F2:F35, "Electronics", D2:D35)	$8,050.00
Affordable Items Count	Number of items priced under $100	=COUNTIF(D2:D35, "<100")	11


Data Transformation & Formula Implementation:

1. String Extraction & Parsing
   
The raw data contained composite Product ID codes (e.g., 03-JUN-CA). Text functions were applied to decompose these strings into actionable attributes: 
•	Day Extraction (First 2 characters):
Excel
=LEFT(A2, 2)
```[cite: 1]
•	Month Extraction (Characters 4 to 3):
Excel
=MID(A2, 4, 3)
```[cite: 1]
•	Country Code Extraction (Last 2 characters):
Excel
=RIGHT(A2, 2)
```[cite: 1]

2. Conditional Categorization
To segment products into premium and budget tiers, a logical IF function was implemented: 
•	Pricing Classification Rule: Items priced above $500 are labeled as High price; all others are labeled as Standard price. 
Excel
=IF(D2>500, "High price", "Standard price")
```[cite: 1]

3. Business Analytics & Aggregation
•	Category Filtering (Electronics Valuation): Evaluated total cost concentration within high-tech inventory. 
Excel
=SUMIF(F2:F35, "Electronics", D2:D35)
```[cite: 1]
•	Budget Segment Analysis: Identified products with entry-level price points (< $100). 
Excel
=COUNTIF(D2:D35, "<100")
```[cite: 1]



Analytical Insights

1.	Category Dominance: Electronics represents $8,050 out of the $10,100 total catalog value (~79.7%), showing high financial dependence on high-tech SKUs. 
2.	Pricing Segmentation: 11 out of 34 items (~32.3%) are budget-friendly (< $100), ensuring accessibility while premium electronics drive overall margin value. 
3.	Data Integrity: Text extraction functions successfully decoupled time and geographic dimensions, enabling downstream regional and seasonal reporting without altering original source tables.
 
Tools Used
•	Tool: Microsoft Excel
•	Techniques: Text Parsing (LEFT, MID, RIGHT), Logical Functions (IF), Conditional Aggregation (SUMIF, COUNTIF), Statistical Analysis (SUM, AVERAGE, MIN, MAX, COUNTA) 




