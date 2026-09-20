# 📊 Excel Data Exploration – Data Analytics

## 📌 Project Overview

This project is part of my **Data Analytics learning journey**, where I explored and analyzed a product dataset using **Microsoft Excel**.

The objective of this project is to build foundational skills in **data exploration, Excel formulas, logical functions, conditional analysis, and text manipulation**.

This project demonstrates how Excel can be used to transform raw product data into meaningful insights through simple analytical techniques.

---

## 🎯 Objectives

* Perform basic data exploration using Excel
* Calculate summary statistics such as **Sum, Count, Average, Minimum, and Maximum**
* Apply logical functions using **IF**
* Perform conditional calculations using **SUMIF and COUNTIF**
* Extract information from text using **LEFT, RIGHT, and MID**
* Create new calculated columns from existing data
* Develop practical Excel skills required for Data Analyst roles

---

## 📁 Dataset

The dataset contains information about different products.

### Dataset Attributes

| Column       | Description                        |
| ------------ | ---------------------------------- |
| Product ID   | Unique identifier for each product |
| Product Name | Name of the product                |
| Brand Name   | Brand associated with the product  |
| Quantity     | Number of products available       |
| Category     | Product category                   |
| Price        | Price of the product               |

---

## 🛠️ Tools & Technologies

* **Microsoft Excel**
* Excel Formulas & Functions
* Data Exploration
* Conditional Analysis
* Text Manipulation

---

## 📋 Tasks Performed

### 1. Basic Data Exploration

Calculated the following:

* **Total Price** of all products
* **Total Number of Products**
* **Average Product Price**

### Excel Functions Used

```excel
=SUM()
=COUNT()
=AVERAGE()
```

---

### 2. Minimum and Maximum Price

Identified the lowest and highest product prices in the dataset.

### Excel Functions Used

```excel
=MIN()
=MAX()
```

---

### 3. Price Range Classification

Created a new column named **Price Range** to categorize products based on their price.

| Condition     | Classification |
| ------------- | -------------- |
| Price >= $500 | High Price     |
| Price < $500  | Standard Price |

### Excel Function Used

```excel
=IF(F2>=500,"High Price","Standard Price")
```

> The cell reference may be changed depending on the location of the Price column in the dataset.

---

### 4. Conditional Analysis

Used conditional Excel functions to analyze products based on category and price.

#### SUMIF

Calculated the **total price of products in the Electronics category**.

```excel
=SUMIF(E:E,"Electronics",F:F)
```

#### COUNTIF

Calculated the **number of products with a price below $100**.

```excel
=COUNTIF(F:F,"<100")
```

---

### 5. Text Extraction

Extracted useful information from the **Product ID** using Excel text functions.

Created the following columns:

| New Column   | Function  | Purpose                    |
| ------------ | --------- | -------------------------- |
| Day          | `LEFT()`  | Extract first 2 characters |
| Country Code | `RIGHT()` | Extract last 2 characters  |
| Month        | `MID()`   | Extract characters 4 to 6  |

### LEFT

```excel
=LEFT(A2,2)
```

Extracts the first two characters from the Product ID.

### RIGHT

```excel
=RIGHT(A2,2)
```

Extracts the last two characters from the Product ID.

### MID

```excel
=MID(A2,4,3)
```

Extracts 3 characters starting from the 4th character of the Product ID.

---

## 📊 Excel Functions Covered

This project helped me practice the following Excel functions:

```text
SUM
COUNT
AVERAGE
MIN
MAX
IF
SUMIF
COUNTIF
LEFT
RIGHT
MID
```

---

## 🔍 Key Learning Outcomes

Through this assignment, I gained practical experience in:

* Understanding and exploring structured datasets
* Performing basic numerical analysis
* Using Excel formulas for data calculations
* Applying logical conditions to classify data
* Performing conditional aggregation
* Extracting specific information from text fields
* Creating new analytical columns
* Preparing raw data for further analysis

---

## 📂 Project Structure

```text
Excel-Data-Exploration/
│
├── Dataset/
│   └── Product_Dataset.xlsx
│
├── Excel/
│   └── Excel_Data_Exploration.xlsx
│
├── Screenshots/
│   └── Excel_Analysis.png
│
└── README.md
```

---

## 🚀 Future Improvements

As I continue my Data Analytics journey, I plan to extend this project by:

* Creating Excel dashboards
* Adding Pivot Tables
* Creating charts and visualizations
* Performing more advanced data cleaning
* Using Power Query
* Learning Power BI
* Performing the same analysis using Python and Pandas
* Building more real-world data analytics projects

---

## 👨‍💻 About Me

I am an **aspiring Data Analyst** currently developing my skills in data analysis, programming, databases, and data visualization.

### Current Skills

* 📊 Microsoft Excel
* 🐍 Python
* 🗄️ SQL
* ☕ Java
* 🌐 HTML
* 🎨 CSS
* ⚡ JavaScript
* 🗃️ MySQL

I am continuously learning and building projects to develop my **Data Analytics portfolio** and gain practical experience with real-world datasets.

---

## 📌 Project Status

**Completed ✅**

This project is part of my **Data Analytics learning portfolio**.

---

⭐ If you find this project useful, feel free to explore my other Data Analytics projects.
