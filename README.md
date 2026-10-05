# Data Cleaning and Transformation Using Power Query & Excel

This project focuses on cleaning, transforming, and formatting a product dataset using **Microsoft Power Query** and **Microsoft Excel**.

The dataset contains missing values, inconsistent text, spelling mistakes, duplicate records, and Product ID information that needs to be extracted and transformed.

The six questions in the assignment were completed using Power Query for data cleaning and transformation and Excel for final formatting and data manipulation.

---
# Dataset
The dataset contains the following fields:
- Product ID
- Product Name
- Brand Name
- Price
- Quantity
- Category
---

# Question 1: Handling Missing Values

## 1.1 Missing Values in the Price Column

### What I Found

The **Price** column contained missing/null values.

### Method Used

I used **Power Query** to identify the missing values and calculated the **median price** to handle the missing price information.

### Steps

1. Loaded the dataset into Power Query.
2. Selected the **Price** column.
3. Checked for null/missing values.
4. Calculated the **median price** using Power Query.
5. Used the median value to handle the missing price values.

### Why Median?

Median was selected because it is less affected by extremely high or low values than the average. It provides a more reliable central value when the dataset contains price variations.

### Result

The missing price information was handled using the calculated median value.

---

## 1.2 Missing Values in the Category Column

### What I Found

Some records contained missing/null values in the **Category** column.

### Method Used

I used **Power Query** to replace the missing category values with:

`Unknown`

### Steps

1. Opened the dataset in Power Query.
2. Selected the **Category** column.
3. Identified the null values.
4. Replaced the null values with `Unknown`.

### Result

All missing category values were replaced with **Unknown**, preventing blank category values in the cleaned dataset.

---

# Question 2: Correcting Inconsistent Data

## 2.1 Inconsistent Product Names

### What I Found

The **Product Name** column contained inconsistent capitalization and spelling.

Examples:

- `laptop` → `Laptop`
- `T-shirt` → `T-Shirt`
- `headphones` → `Headphones`

### Method Used

I used **Excel Find and Replace** to correct the inconsistent Product Name values.

### Steps

1. Selected the **Product Name** column in Excel.
2. Pressed **Ctrl + H** to open Find and Replace.
3. Replaced the incorrect values with the standardized values.
4. Checked the column to ensure consistent naming.

### Result

The Product Name values were standardized.

---

## 2.2 Cleaning Product Names

### Method Used

I used **Power Query** to further clean the Product Name column.

### Steps

1. Selected the **Product Name** column.
2. Used **Transform → Format → Trim**.
3. Used **Transform → Format → Clean**.

### Purpose

- **Trim** removes unnecessary spaces.
- **Clean** removes non-printable or unwanted characters.

### Result

The Product Name column was cleaned and standardized.

---

## 2.3 Correcting Category Spelling Mistakes

### What I Found

A spelling mistake was found in the Category column:

`Electrioni`

The correct value is:

`Electronics`

### Method Used

I used a **Custom Column with conditional logic in Power Query**.

### Steps

1. Opened **Add Column → Custom Column**.
2. Created a condition to check the Category value.
3. If the value was `Electrioni`, it was changed to `Electronics`.
4. Otherwise, the original category was retained.

### Result

`Electrioni` was corrected to `Electronics`.

---

# Question 3: Removing Duplicate Records

### What I Found

The dataset contained duplicate rows.

Examples included:

- **Laptop Bag – Samsonite**
- **Laptop – HP**

### Method Used

I used **Power Query → Remove Rows → Remove Duplicates**.

### Steps

1. Loaded the dataset into Power Query.
2. Selected the columns representing the complete record.
3. Used **Home → Remove Rows → Remove Duplicates**.
4. Power Query removed duplicate records while keeping the first occurrence.

### Result

Duplicate records were removed from the dataset.

---

# Question 4: Splitting and Merging Data

## 4.1 Splitting Product ID

### What I Found

The Product ID contained multiple pieces of information.

Examples:

```text
28-JAN-US
15-FEB-US
03-MAR-US

The information needed to be separated into:

Manufacturing Date

Country Code

Method Used
I used Power Query Extract and Delimiter functions.

Steps
Selected the Product ID column.

Used the - delimiter.

Extracted the required text before and after the delimiter.

Separated the information into appropriate columns.

Created the Manufacturing Date and Country Code fields.

Result
The Product ID information was separated into meaningful fields.

Example:

Manufacturing Date	Country Code
28-01-2026	US
15-02-2026	US
03-03-2026	US

4.2 Creating the Manufacturing Date
Method Used
I used Power Query to create the Manufacturing Date from the extracted Product ID information.

Steps
Used the extracted date components.

Created the Manufacturing Date using Power Query.

Changed the column data type to Date.

Loaded the transformed data back into Excel.

Result
The Manufacturing Date was stored as a proper Date value instead of text.

4.3 Merging Product Name and Brand Name
What I Needed to Do
The assignment required the Product Name and Brand Name to be combined into a new column called:

Product Brand

Method Used
I used Microsoft Excel and the & operator.

Steps
Created a new column named Product Brand.

Used the formula:

=B2&" "&C2

Filled the formula down for all records.

Example
Laptop + Dell → Laptop Dell

Result
A new Product Brand column was created containing the combined Product Name and Brand Name.

Question 5: Number Formatting
5.1 Formatting Price as Currency
Method Used
I used Microsoft Excel to format the Price column as currency.

Steps
Selected the Price column.

Opened the Number Format options.

Selected Currency.

Applied the Indian Rupee (₹) format.

Example
1000 → ₹1,000.00
950  → ₹950.00
130  → ₹130.00

Result
The Price column was displayed in a consistent currency format.

5.2 Formatting Manufacturing Date
Method Used
I used Power Query to convert the Manufacturing Date into a proper Date data type.

After loading the data into Excel, the date was displayed in:

DD-MM-YYYY

Example
28-01-2026
05-02-2026
03-03-2026

Result
The Manufacturing Date was displayed consistently in the required format.

Question 6: Conditional Formatting
6.1 Conditional Formatting for Price
Method Used
I used Microsoft Excel Conditional Formatting on the Price column.

Steps
Selected the Price column.

Went to Home → Conditional Formatting.

Applied a Data Bar or Color Scale.

Result
The prices became easier to compare visually.

Higher prices are visually distinguished from lower prices using the selected formatting.

6.2 Highlighting Electronics
Method Used
I used Microsoft Excel Conditional Formatting to highlight the Electronics category.

Steps
Selected the Category column.

Went to Home → Conditional Formatting → New Rule.

Created a rule for:

Electronics

Selected a highlight/fill color.

Applied the rule.

Result
All cells containing Electronics were highlighted.

Tools and Techniques Used
Question	Task	Tool / Technique
Q1(a)	Missing Price	Power Query – Median
Q1(b)	Missing Category	Power Query – Replace Null with Unknown
Q2(a)	Product Name correction	Excel – Find & Replace
Q2(b)	Category spelling correction	Power Query – Custom Column & Conditional Logic
Q2(c)	Remove extra spaces	Power Query – Trim
Q2(c)	Remove unwanted characters	Power Query – Clean
Q3	Remove duplicate records	Power Query – Remove Duplicates
Q4(a)	Split Product ID	Power Query – Extract & Delimiter
Q4(a)	Country Code	Power Query – Extract
Q4(a)	Manufacturing Date	Power Query
Q4(b)	Merge Product + Brand	Excel – & Operator
Q5(a)	Price formatting	Excel – Currency Format
Q5(b)	Date formatting	Power Query + Excel
Q6(a)	Price visualization	Excel – Conditional Formatting
Q6(b)	Highlight Electronics	Excel – Conditional Formatting

Final Outcome
The dataset was successfully cleaned and transformed by:

Handling missing Price values using the median.

Replacing missing Category values with Unknown.

Correcting inconsistent Product Names using Excel Find & Replace.

Cleaning text using Power Query Trim and Clean.

Correcting Category spelling mistakes using Power Query conditional logic.

Removing duplicate records using Power Query.

Extracting information from Product ID using delimiters.

Creating the Manufacturing Date and Country Code.

Merging Product Name and Brand Name using Excel.

Formatting Price as Indian currency.

Converting Manufacturing Date to a proper date.

Applying conditional formatting to Price and Category.

The final dataset is clean, consistent, structured, and ready for further analysis.

Tools Used
Microsoft Power Query
Microsoft Excel
