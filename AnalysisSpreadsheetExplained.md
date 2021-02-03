/*

This script analyzes Israeli credit cards 
and builds a db people can build nice dashboards from.
Instructions for a new Sheet:
1. make sure the new spreadsheet has the following tabs:
- DB
- Categories
- _status (a utility tab to summarize some data for control purposes)
- _categories_flattened (a utility tab to hold temporary calculations)
- _conversion_rates (a utility tab to hold temporary calculations)

2. Inside DB sheet (tab) - make sure the following headers appear on the first row:
- Date Added
- File ID	
- File Name	
- Card Type
- Card Number
- Billing Date
- Transaction Date
- Business Name
- Amount
- Currency
- Category
-- this column is created automatically by using the following formula:
"=arrayformula(iferror(vlookup(F2:F,'_categories_flattened'!A2:B,2,FALSE),"ללא סיווג"))"
- Conversion Rate	
-- this column is created automatically by using the following formula:
"=arrayformula(if(J2:J="ILS",1,IF(J2:J="USD",VLOOKUP(G2:G+1,'_conversion_rates'!A4:B,2,TRUE),IF(J2:J="EUR",VLOOKUP(G2:G+1,'_conversion_rates'!C4:D,2,TRUE),"ERR"))))"
- Amount ILS
-- this column is created automatically by using the following formula:
"=arrayformula(I2:I*L2:L)"

3. Categories sheet 
- This tab is handled manually (you can choose your own categories and where each business goes)
- The only mandatory column is the last one which is "ללא סיווג". Its first row should run the following formula:
"=QUERY(DB!F2:I,"select F where I='ללא סיווג'",0)"

4. Inside _status sheet - make sure the following headers appear:
-File Name
-File ID

5. _categories_flattened sheet
This tab organizes the categories table in a more processing-efficient way.
Headers:
- Business name
-- "=filter(flatten(Categories!A2:O),NOT(ISBLANK(flatten(Categories!A2:O))))"
- Category
-- "=arrayformula(reverseLookup(A2:A,transpose(Categories!A:O)))"

6. _conversion_rates sheet
used formulas:
A1: "=min(DB!G:G)"
A2: "=MAX(DB!G:G)"
A5: "=query(GOOGLEFINANCE("CURRENCY:USDILS","price",A1, A2, "DAILY"),"select * label Col1 '', Col2''")"
C5: "=query(GOOGLEFINANCE("CURRENCY:EURILS","price",A1, A2, "DAILY"),"select * label Col1 '', Col2''")"

*/