# -DATA-EXPLORATION
MY FULL EXPLANATION OF DATA EXPLORATION
Dataset Overview & SchemaThe dataset consists of 34 unique product entries across categories such as Electronics, Fashion, Kitchen, Outdoor, and Accessories.   Field NameDescriptionExample ValueProduct IDComposite string containing metadata03-JUN-CA   Product NameName of the product itemLAPTOP, SUNGLASSES   Brand NameManufacturer / BrandDELL, RAY-BAN   Price ($)Unit price of the product$1,000, $130   QuantityStock / Order Quantity30, 25   CategoryProduct sectorELECTRONICS, FASHION   Price RangeDerived pricing tier (High Price vs Standard Price)HIGH PRICE, STANDARD PRICE   DayExtracted day from Product ID3, 28   Country CodeExtracted country destination from Product IDCA, US   MonthExtracted transaction month from Product IDJUN, JAN   Key Performance Indicators (KPIs)MetricBusiness InsightExcel FormulaValueTotal ProductsTotal volume of tracked SKUs=COUNTA(B2:B35)   34   Total Inventory ValueCumulative sum of single unit prices=SUM(D2:D35)   $10,100   Average Unit PriceMean product price across catalog=AVERAGE(D2:D35)   $297.06   Minimum PriceLowest entry-level item price=MIN(D2:D35)   $30.00   Maximum PriceHighest premium item price=MAX(D2:D35)   $1,000.00   Electronics Category ValueTotal price value of Electronics items=SUMIF(F2:F35, "Electronics", D2:D35)   $8,050.00   Affordable Items CountNumber of items priced under $100=COUNTIF(D2:D35, "<100")   11   Data Transformation & Formula Implementation1. String Extraction & ParsingThe raw data contained composite Product ID codes (e.g., 03-JUN-CA). Text functions were applied to decompose these strings into actionable attributes:   Day Extraction (First 2 characters):Excel=LEFT(A2, 2)
```[cite: 1]
Month Extraction (Characters 4 to 6):Excel=MID(A2, 4, 3)
```[cite: 1]
Country Code Extraction (Last 2 characters):Excel=RIGHT(A2, 2)
```[cite: 1]

2. Conditional CategorizationTo segment products into premium and budget tiers, a logical IF function was implemented:   Pricing Classification Rule: Items priced above $500 are labeled as High price; all others are labeled as Standard price.   Excel=IF(D2>500, "High price", "Standard price")
```[cite: 1]

3. Business Analytics & AggregationCategory Filtering (Electronics Valuation): Evaluated total cost concentration within high-tech inventory.   Excel=SUMIF(F2:F35, "Electronics", D2:D35)
```[cite: 1]
Budget Segment Analysis: Identified products with entry-level price points (< $100).   Excel=COUNTIF(D2:D35, "<100")
```[cite: 1]

Analytical InsightsCategory Dominance: Electronics represents $8,050 out of the $10,100 total catalog value (~79.7%), showing high financial dependence on high-tech SKUs.   Pricing Segmentation: 11 out of 34 items (~32.3%) are budget-friendly (< $100), ensuring accessibility while premium electronics drive overall margin value.   Data Integrity: Text extraction functions successfully decoupled time and geographic dimensions, enabling downstream regional and seasonal reporting without altering original source tables.   Tools UsedTool: Microsoft ExcelTechniques: Text Parsing (LEFT, MID, RIGHT), Logical Functions (IF), Conditional Aggregation (SUMIF, COUNTIF), Statistical Analysis (SUM, AVERAGE, MIN, MAX, COUNTA)   
