# Product and Branch Performance Analysis 📈

This project analyzes product and branch performance based on monthly sales data, generating visual plots and exporting insights to an Excel file.

## Features
- Cleans and preprocesses sales data
- Calculates revenue, profit, and margin
- Identifies best and worst products/branches/months
- Generates:
  - Line and bar plots
  - Written summaries in Excel
  - Embedded images and data tables

## How It Works
1. Load and clean sales data from a `.csv` file
2. Map full month names to short ones and order them
3. Group data by product, branch, and month
4. Save:
   - Data tables
   - Plots
   - Textual reports
   to different Excel sheets
5. Visualizations are created using `matplotlib` and `seaborn`

## Output
- 📊 Excel file `product.xlsx` with:
  - Tables: revenue & profit per product/branch/month
  - Reports: analysis written into separate sheets
  - Plots: added as images to Excel

## Requirements
- Python 3.8+
- pandas
- numpy
- matplotlib
- seaborn
- openpyxl

## License
This project is licensed under the MIT License.
