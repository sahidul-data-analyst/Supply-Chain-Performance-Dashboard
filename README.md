# Supply Chain Performance Dashboard (MIS)

### Project Overview
This dashboard was developed to provide actionable insights into supply chain operations, focusing on production efficiency, quality control, and cost management. As an MIS professional with 3.5 years of experience, I designed this tool to monitor key metrics that directly impact operational success.

### Key Metrics Tracked(Power BI)
**Total Shipping Cost**
Total Shipping Cost=SUM('Supply Chaine Performance Data'[Shipping Cost ($)])
**Total Unique Supplier**
Total Unique Supplier= DISTINCTCOUNT('Supply Chaine Performance Data'[Supplier Name])
**Total Defect Rate %**
Total Defect Rate %= DIVIDE(SUM('Supply Chaine Performance Data'[Defective Units]),SUM('Supply Chaine Performance Data'[Total Unit]))
**Average Lead Time**
Average Lead Time= AVERAGE('Supply Chaine Performance Data'[Lead Time])
**High Risk Supplier**
High Risk Supplier= CALCULATE(SUM('Supply Chaine Performance Data'[Shipping Cost ($)]),'Supply Chaine Performance Data'[Defect Rate Status]="High Risk")
**Delivery Performance**
Delivery Performance= SWITCH(TRUE(),'Supply Chaine Performance Data'[Lead Time]<=3,"Express",'Supply Chaine Performance Data'[Lead Time]<=8,"Normal Delivery","Delayed")
**LY Shipping Cost $**
LY Shipping Cost $= CALCULATE([Total Shipping Cost],SAMEPERIODLASTYEAR('Date Table'[Date]))
**Date Table**
Date Table= CALENDAR(DATE(2010,1,1),DATE(2026,12,31))
**Month Name**
Month Name= FORMAT('Date Table'[Date],"MMMM")
**Months Number**
Months Number= MONTH('Date Table'[Date])
**Year**
Year= FORMAT('Date Table'[Date],"YYYY")
**Shipping Cost Growth %**
Shipping Cost Growth %= DIVIDE([Total Shipping Cost]-[LY Shipping Cost $][LY Shipping Cost $])
**Total Shipping Cost MTD**
Total Shipping Cost MTD= TOTALMTD([Total Shipping Cost],'Date Table'[Date])
**Total Shipping Cost YTD**
Total Shipping Cost YTD= TOTALYTD([Total Shipping Cost],'Date Table'[Date])
**Running Total Shipping Cost $**
Running Total Shipping Cost $= CALCULATE([Total Shipping Cost],FILTER(ALLSELECTED('Date Table'),'Date Table'[Date]<=MAX('Date Table'[Date])))
**Supplier Rank**
Supplier Rank= RANKX(ALL('Supply Chaine Performance Data'[Supplier Name]),'Supply Chaine Performance Data'[Total Defect Rate %], ,DESC,Dense)
**Delivey Shipping Cost**
Delivey Shipping Cost= CALCULATE(SUM('Supply Chaine Performance Data'[Shipping Cost ($)]),USERELATIONSHIP('Date Table'[Date],'Supply Chaine Performance Data'[Delivery Date]))
**Defect Rate By Delimetar**
Defect Rate By Delimetar= CALCULATE(SUM('Supply Chaine Performance Data'[Defect Rate]),USERELATIONSHIP('Date Table'[Date],'Supply Chaine Performance Data'[Delivery Date])),
### Key Metrics Tracked(Excel 50+ Functions)
**IF/Switch**
Defect Rate Status=IFS([Defect Rate]>=10%, "High Risk", [Defect Rate]>=3%, "Medium Risk", TRUE, "Low Risk")
**index-match**
=INDEX(Table1[[#All],[Defective Units]],MATCH(J4,Table1[[#All],[Production Line Id]],0))
**Xlookup**
=XLOOKUP("SL-101", A:A, F:F)
**Sumif/Countif/Averageif**
=SUMIF(Table1[[#All],[Defect Type]],C4,Table1[[#All],[Defective Units]])
=COUNTIF(Table1[[#All],[Defect Type]],J39)
=AVERAGEIF(Table1[[#All],[Defect Type]],J26,Table1[[#All],[Defective Units]])
**Offset**
=OFFSET($A$4, 0, 0, COUNTA($A:$A),8)
**Filter/Unique/Short**
=UNIQUE(C4:C44)
=FILTER(A4:H44, G4:G44="High Risk")
=SORT(A4:H44, 5, -1)
**Aggregate/Subtotal/AND/OR**
=SUBTOTAL(9, E4:E44)
=AGGREGATE(1, 6, H4:H44)
=IF(AND(F4>5%, E4>500), "Critical", "Normal")
=IF(OR(C4="Broken Stitch", C4="Oil Stain"), "Re-check", "Pass")
### Key Metrics Tracked(SQL)
**Replace/update function**
UPDATE[Supply Chain Performance Data.xlsx] 
SET[Shipping Cost]=REPLACE([Shipping Cost], '$','');
**ISNUMERIC**
SELECT[Defect Rate]
FROM[Supply Chain Performance Data.xlsx]
WHERE ISNUMERIC([Defect Rate])=0;
**ISDATE**
SELECT[Delivery Date]
FROM[Supply Chain Performance Data.xlsx]
WHERE ISDATE ([Delivery Date])=0;
**TRY_CAST**
SELECT[Order Date],[Delivery Date]
FROM[Supply Chain Performance Data.xlsx]
WHERE TRY_CAST([Order Date]AS DATE)IS NULL
OR TRY_CAST(Delivery Date]AS DATE)IS NULL;
**Alter**
ALTER Table[Supply Chain Performance Data.xlsx]
ALTER COLUMN[Defect Rate]FLOAT;
**CASE WHEN (Automated Risk Status):**
SELECT [Supplier Name], [Defect Rate],
  CASE 
      WHEN [Defect Rate] > 5 THEN 'High Risk'
      WHEN [Defect Rate] BETWEEN 3 AND 5 THEN 'Medium Risk'
      ELSE 'Low Risk'
  END AS RiskStatus
  FROM [dbo].[Supply Chain Parformance Data];
  **SQL Joins(Merging Data Tables)**
  SELECT Orders.OrderID, Orders.OrderDate, Suppliers.SupplierName
FROM Orders
INNER JOIN Suppliers ON Orders.SupplierID = Suppliers.SupplierID;
**DATEDIFF(Lead Time Calculation)**
SELECT [Supplier Name], [Order Date], [Delivery Date],
DATEDIFF(day, [Order Date], [Delivery Date]) AS LeadTimeDays
FROM [dbo].[Supply Chain Performance Data];
**Grouping & Aggregations(Summarized Reporting)**
SELECT [Supplier Name], 
       COUNT([Order ID]) AS TotalOrders,
       SUM([Total Units]) AS TotalProduction,
       AVG([Defect Rate]) AS AvgDefectRate
FROM [dbo].[Supply Chain Performance Data]
GROUP BY [Supplier Name];
**COALESCE**
SELECT
     COALESCE(SupplierName, 
     'Unknown') FROM Suppliers;
**HAVING**
SELECT SupplierID,
COUNT(OrderID) FROM Orders 
GROUP BY SupplierID HAVING 
COUNT(OrderID) > 10;
### Technical Skills Demonstrated
1. Advanced Data Automation (Excel Expertise - 50+ Functions):

Dynamic Lookup & Reference: Expert in XLOOKUP, INDEX-MATCH, OFFSET, and INDIRECT for building automated reporting systems.

Array Formulas: Utilizing FILTER, UNIQUE, SORT, and SEQUENCE to create dynamic, self-updating dashboards.

Precision Reporting: Leveraged AGGREGATE, SUBTOTAL, and IFERROR to ensure 100% accuracy in large industrial datasets.

ETL Processes: Proficient in Power Query for advanced data cleaning, merging, and transformation.

2. Business Intelligence & Analytics (Power BI & DAX):

Advanced DAX Modeling: Skilled in complex measures using CALCULATE, ALLSELECTED, FILTER, and KEEPFILTERS.

Time Intelligence: Implementation of SAMEPERIODLASTYEAR, DATEADD, TOTALMTD, and TOTALYTD for YoY and MoM performance tracking.

KPI Development: Created performance metrics using RANKX, DIVIDE, SWITCH, and RELATED to monitor production efficiency.

Context Control: Expertise in managing filter context through proper relationship modeling and USERELATIONSHIP.

3. Database Management & SQL:

Industrial Querying: Proficient in SELECT, CASE WHEN logic, and complex JOINS (Inner, Left, Right) to extract insights from ERP databases.

Data Aggregation: Skilled in using GROUP BY, HAVING, and COUNT/SUM to summarize supply chain operational data.

Advanced Formatting: Expert in DATEDIFF and DATE_FORMAT for calculating lead times and logistics cycles.

4. Supply Chain & MIS Operations:

Operational Excellence: Track record of using data to identify bottlenecks, resulting in a 10% reduction in wastage.


### Business Impact
By identifying top suppliers with high defect rates and monitoring shipping cost trends, this dashboard facilitates data-driven decisions that help in reducing wastage and optimizing lead times.
