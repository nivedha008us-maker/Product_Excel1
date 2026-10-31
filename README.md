# Product_Excel1
Data_exploration_assignment
📊 Product Dataset – Excel Data Analysis Project
📌 Project Overview

This project focuses on analyzing a product dataset using Microsoft Excel. It demonstrates fundamental Excel functions, logical functions, conditional calculations, text manipulation, and data formatting techniques.

The dataset contains 34 product records from different brands, categories, and countries. The project uses Excel formulas to summarize product prices, identify minimum and maximum values, perform category-based calculations, and classify products based on their prices.

📁 Dataset Description

The dataset contains information about products, including their names, brands, prices, quantities, categories, and country details.

Dataset Columns


Column Name	Description
Product ID	Unique identifier for each product
Product Name	Name of the product
Brand Name	Brand associated with the product
Price ($)	Price of the product in dollars
Quantity	Available quantity of the product
Category	Product category
Price Range (Price < 500)	Price classification using IF
Price Range (Price >= 500)	Price classification using IF
Day	Day extracted from the Product ID
Country Code	Country code extracted from the Product ID
📈 Dataset Information

Total Records: 34
Total Price: $10,100
Average Price: $297.06
Minimum Price: $30
Maximum Price: $1,000
Number of Products: 34
🛠️ Tools & Technologies

Microsoft Excel
Excel Formulas and Functions
Data Cleaning and Formatting
Conditional Formatting
Data Analysis
🧮 Excel Functions Used

1. SUM, COUNT, and AVERAGE

These functions are used to calculate basic statistical information about the products.


Function	Purpose	Result
SUM()	Calculate the total product price	$10,100
COUNT()	Count the number of products	34
AVERAGE()	Calculate the average product price	$297.06
Example formulas:


excel
=SUM(TAB_DB[Price ($)])
=COUNT(TAB_DB[Price ($)])
=AVERAGE(TAB_DB[Price ($)])
2. MIN and MAX

These functions identify the lowest and highest product prices.


Function	Purpose	Result
MIN()	Find the minimum product price	$30
MAX()	Find the maximum product price	$1,000
Example formulas:


excel
=MIN(TAB_DB[Price ($)])
=MAX(TAB_DB[Price ($)])
3. SUMIF and COUNTIF

These functions perform calculations based on specific conditions.

SUMIF: Calculates the total price of products belonging to the Electronics category.


excel
=SUMIF(TAB_DB[Category],"Electronics",TAB_DB[Price ($)])
Result: $8,050

COUNTIF: Counts the number of products with a price below $100.


excel
=COUNTIF(TAB_DB[Price ($)],"<100")
Result: 11 products

4. IF Function – Logical Classification

The IF function classifies products based on their prices.

Conditions:

If Price >= $500 → High Price
If Price < $500 → Standard Price
Excel formula:


excel
=IF([@[Price ($)]]>=500,"High price","Standard price")
This creates a new column called Price Range.

5. Text Functions – LEFT, RIGHT, and MID

Text functions are used to extract specific information from the Product ID column.

LEFT() – Extracts characters from the beginning of the Product ID.
RIGHT() – Extracts characters from the end of the Product ID.
MID() – Extracts characters from a specified position within the Product ID.
Example formulas:


excel
=LEFT([@[Product ID]],2)
=RIGHT([@[Product ID]],2)
=MID([@[Product ID]],3,2)
These functions can be used to create additional columns such as Day and Country Code, depending on the Product ID structure.

📊 Key Analysis Results


Analysis	Result
Total Product Price	$10,100
Total Products	34
Average Product Price	$297.06
Minimum Product Price	$30
Maximum Product Price	$1,000
Electronics Category Total	$8,050
Products Below $100	11
🎯 Project Objectives

Calculate the total price of all products.
Determine the total number of products.
Calculate the average product price.
Identify the minimum and maximum product prices.
Calculate category-based totals using SUMIF.
Count products based on price conditions using COUNTIF.
Classify products using the IF function.
Extract information using LEFT, RIGHT, and MID functions.
Improve data organization and presentation using Excel formatting.
📚 Skills Demonstrated

Data Analysis
Excel Functions and Formulas
Logical Functions
Conditional Calculations
Text Manipulation
Structured Table References
Data Formatting
Problem Solving
👨‍💻 Conclusion

This project demonstrates how Microsoft Excel can be used to analyze and organize product-related data efficiently. By applying different Excel functions, meaningful insights can be extracted from raw data.

The project helped develop practical knowledge of Excel formulas, conditional calculations, logical functions, and text manipulation, which are essential skills for data analysis.

Project Type: Data Analysis
Tool: Microsoft Excel
Dataset Size: 34 Records
Domain: Product and Retail Data Analysis
