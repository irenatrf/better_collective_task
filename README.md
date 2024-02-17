# Order Return Analysis Assignment for Better Collective

This `main.py` script performs an analysis of order returns based on order and return datasets for the years 2015 to 2018.

## Instructions

1. Ensure that the required libraries are installed:

   ```bash
   pip install pandas pygsheets
   ```

2. Place the data files (`Orders_Central.csv`, `orders_south_2015.csv`, `return reasons_new.xlsx`) in the `datasets` directory before running the script.

3. Run the script to perform the analysis.

## Script Overview

1. **Data Loading:** Load the order and return reason datasets into pandas DataFrames.
2. **Data Processing:** Process and clean the order data for each year (2015, 2016, 2017, 2018).
3. **Central Orders Processing:** Clean and filter the central orders dataset.
4. **Invalid  Dates Check:** Check for invalid  dates where the order year is greater than the ship year.
5. **Data Merging:** Merge the cleaned order datasets into a single DataFrame.
6. **Return Reasons Processing:** Process the return reasons dataset, filter by year, and merge with the order data.
7. **Return Rate Analysis:** Calculate the return rate (%) per year.
8. **Product Category Analysis:** Analyze the most returned product categories.

## Output

- The output files that are stored in output files directory are also stored in Google Sheets for analytics visualization in Looker Studio

Google Sheets 
https://docs.google.com/spreadsheets/d/1fjMzHA_2S9X9I0ll5rPDwpzzN-i_e0sHIsIK-uZfP4Y/edit?usp=sharing

Looker Studio
https://lookerstudio.google.com/reporting/f7235198-f2f2-4ab6-b1e1-323be194f8b5
