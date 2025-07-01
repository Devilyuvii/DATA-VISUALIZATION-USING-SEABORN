# DATA-VISUALIZATION-USING-SEABORN

from pathlib import Path

## Project Title
**Excel Data Cleaner and Visualizer using Python**

---

## Overview

This Python project automates the process of working with multiple Excel files. It reads Excel sheets from a folder, removes duplicates, combines them into a single clean dataset, saves the result to a new Excel file, and creates basic visualizations (line plot, scatter plot, heatmap) for two key columns (`a` and `b`).

This tool is ideal for analysts and data engineers who frequently deal with Excel files and need a quick, repeatable way to preprocess and explore data.

---

## Steps Involved in the Project

1. **Read** all `.xls` and `.xlsx` files from a given folder
2. **Remove** duplicates from each file and across the combined dataset
3. **Merge** all cleaned data into a single DataFrame
4. **Export** the final cleaned data to a specified Excel file
5. **Visualize** the relationship between two columns (`a` and `b`) through:
   - Line Plot
   - Scatter Plot
   - Correlation Heatmap

---

## Sample Dataset Format

Your input Excel files should contain the following columns at minimum:

| a     | b     | other_columns... |
|-------|-------|------------------|
| 10    | 5.6   | ...              |
| 15    | 7.1   | ...              |
| 10    | 5.6   | ...              |
| 20    | 9.3   | ...              |

> Column `a` and `b` are used for visualization. You can rename or change them in the script if your actual data uses different headers.

---

## Features

-  Supports both `.xls` and `.xlsx` formats
-  Duplicate row removal (per file and globally)
-  Combined dataset export to Excel
-  Built-in plotting (line, scatter, heatmap)
-  Simple CLI-based user inputs
-  Easy to customize for different columns

---

## Sample Use Cases

- Combine monthly sales Excel files and visualize trends
- Clean repeated form submission data from multiple `.xls` files
- Analyze correlation between any two numeric variables in survey results
- Data pre-processing before feeding into machine learning models

---

## How to Run the Project

1. **Install Required Libraries**

pip install pandas seaborn matplotlib openpyxl

---

# Run the Script

python excel_cleaner_visualizer.py

---

# Future Improvements

Add GUI support (Tkinter or PyQt)

Support for column selection during runtime

Enhanced error handling and data validation

Export plots to image files

Summary statistics generation

Integration with cloud storage (e.g., Google Drive, S3)

---

# Conclusion & Results

This project automates a tedious part of Excel data cleaning and initial visualization. By reducing manual errors and standardizing the process, it saves time and provides immediate insight through visual tools. Ideal for rapid data prototyping and reporting tasks.

---
