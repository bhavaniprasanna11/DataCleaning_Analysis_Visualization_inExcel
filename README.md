# “Data Cleaning, Analysis & Visualization with Excel” 


# 📦 E-commerce store dataset

## 📌 Project Overview

This dataset contains Four Tables Product Table having Product details, Customer Table having Customer information, Store Table having Store information and Sales Table having Sales information.

It can be used for practicing data analysis, data visualization, Excel, SQL, Python, and Power BI.


## 🎯 Project Objectives

- Data entry, import, and basic organization using tables
- Data cleaning: find/remove duplicates, handle missing/inconsistent data, correct formatting errors, and validate data
- Remove the column that contains irrelevant data if it exists
- Application of Excel formulas for calculations, summaries, and logical/text operations
- Use the XLOOKUP/VLOOKUP functions to combine the tables.
- Use of Pivot Tables for advanced summarization
- Descriptive statistics (mean, median, mode, standard deviation, etc.) using formulas or Analysis ToolPak
- Data visualization: create at least three chart types, build an interactive dashboard, and use basic PivotCharts



## 🛠️ Tools and Technologies

This dataset can be used with:

- Microsoft Excel

  📂 Dataset Description
The dataset contains the following columns:
<img width="837" height="182" alt="image" src="https://github.com/user-attachments/assets/a69f7efc-5a12-432e-be8d-371496558455" />
<img width="692" height="192" alt="image" src="https://github.com/user-attachments/assets/9aa8c7cd-ed9e-41f1-bd4b-a4e404008cc0" />
<img width="977" height="220" alt="image" src="https://github.com/user-attachments/assets/e736d247-7971-4d37-80ee-a4f7c68166b8" />

###**Data Cleaning:-**

1. Customer Table - 	Checking for blanks cells in Loyalty level column and to fill the datas using **if** function and **isblank** function. 

Formula Used:-

-**=IF(ISBLANK([@[Loyalty_Level]]),"Platinum",[Loyalty_Level])**
-**=PROPER(TRIM(B2))**

<img width="1287" height="582" alt="image" src="https://github.com/user-attachments/assets/c0e63952-8e11-410a-bb5e-fd5bf89095cd" />

2. Product Table - Checking the blank cells in Cost and Stock column used **round, if, isblank and Average** to fill the missing data and for Stock column used **if, isblank** function.

Formula Used:-

-**=ROUND(IF(ISBLANK([@Cost]),AVERAGE([Cost]),[@Cost]),0)**
-**=IF(ISBLANK([@Stock]),"0",[@Stock])**

<img width="1540" height="537" alt="image" src="https://github.com/user-attachments/assets/a43a17f5-8406-4ce7-b93c-8b6c2ad06521" />

3. Sales Table - In the sales table to fill the blank cells in the Column Quantity, Discount, Unit price and Total Amount i used below formulas.

Formula Used:-

-**=IF(ISBLANK([@Quantity]),"Unknown",[@Quantity])**
-**=ROUND(IF(ISBLANK([@[Unit_Price]]),AVERAGE([Unit_Price]),[@[Unit_Price]]),0)**
-**=IF(ISBLANK([@Discount]),"0",[@Discount])**
-**=IF(ISBLANK([@[Total_Amount]]),"Unknown",[@[Total_Amount]])**

<img width="1887" height="432" alt="image" src="https://github.com/user-attachments/assets/51792550-d4a1-4a6d-8564-2dc2106ca0da" />

###**Remove Duplicates:-**

- Using Power query for sales table to remove the duplicate rows.

<img width="1682" height="391" alt="image" src="https://github.com/user-attachments/assets/6898c6b7-71d8-4572-afc5-bdaafe7c9e3b" />

###**Vlookup Functions:-**

- To combine customer, product and Store table in a Sales table I used Vlookup functions.

Formula Used:-

-**=VLOOKUP([@[Customer_ID]],Customer_DB[[#All],[Customer_ID]:[State]],7,)**
-**=VLOOKUP([@[Product_ID]],Product_DB[#All],9,)**
-**=VLOOKUP([@[Store_ID]],Store_DB[#All],3,)**
-**=VLOOKUP([@[Customer_ID]],Customer_DB[[#All],[Customer_ID]:[Gender]],5,)**

<img width="1872" height="396" alt="image" src="https://github.com/user-attachments/assets/5047fd72-d762-4e57-b84d-c5efa9e54002" />

###**Pivot Tables:-**

- Created Pivot Tables for advanced summarization functions like finding the Total sales Region wise and State wise, Gender wise Average sales and Count. Top 5 States Sales.

<img width="486" height="165" alt="image" src="https://github.com/user-attachments/assets/88309d5c-8d82-48e8-a5f5-52876e20d338" />

<img width="480" height="147" alt="image" src="https://github.com/user-attachments/assets/4d6f2a72-19f6-4b94-9863-710d69392754" />

<img width="302" height="170" alt="image" src="https://github.com/user-attachments/assets/8119d626-74ef-4707-babf-3e41249dc89e" />

<img width="307" height="557" alt="image" src="https://github.com/user-attachments/assets/a761e129-ca46-496a-a971-ddb6fe6e255a" />

###**Descriptive Statistics:-**

Formula Used:-

-**=SUM(Sales_DB[Column1])**
-**=MIN(Sales_DB[Column1])**
-**=MAX(Sales_DB[Column1])**
-**=AVERAGE(Sales_DB[Column1])**
-**=MEDIAN(Sales_DB[Column1]**)
-**=COUNTIF(Sales_DB[Column1],">=500")**
-**=STDEV.S(Sales_DB[Column1])**

<img width="476" height="200" alt="image" src="https://github.com/user-attachments/assets/f63f3224-8b1d-4375-9303-0d6610152fad" />

###**Data Visualization:-**

- Created the Dash board sheet in Excel and used four types Pivot charts for Dash board designing.

<img width="1747" height="727" alt="image" src="https://github.com/user-attachments/assets/43ae6852-7b7b-4210-b932-6e1e3111927b" />


##**Key Learnings**
Through this project, I practiced:

- statistical analysis
- Mathematical and Aggregate functions.
- Text functions.
- Power Query using Remove duplicates.
- Data imputation using ISBLANK in Numerical and Categorical.
- Data Cleaning and Analysis
- Pivot Tables and Charts
- Dash board Designing


## 👤 Author

**Bhavani**

Aspiring Data Analyst | Excel | SQL | Python | Power BI

---
