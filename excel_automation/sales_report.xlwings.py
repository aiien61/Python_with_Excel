import pandas as pd
import xlwings as xw
from pathlib import Path
from typing import List


# Directory of this file
this_dir = Path(__file__).resolve().parent.parent

# Read in all files
parts: List[pd.DataFrame] = []
for path in (this_dir / "data/sales_data").rglob("*.xls*"):
    print(f"Reading {path.name}")
    part = pd.read_excel(path)
    parts.append(part)

# Combine the DataFrames from each file into a single DataFrame
df = pd.concat(parts)

# Pivot each store into a column and sum up all transactions per date
pivot = pd.pivot_table(df, 
                       index='transaction_date', 
                       columns='store', 
                       values='amount', 
                       aggfunc='sum')

# Resample to end of month and assign an index name
summary = pivot.resample('M').sum()
summary.index.name = 'Month'

# Sort columns by total revenue
summary = summary.loc[:, summary.sum().sort_values().index]

# Add row column totals: Using 'append' together with 'rename' 
# is a convenient way to add a row to the bottom of a DataFrame
summary.loc[:, 'Total'] = summary.sum(axis=1)
summary = pd.concat([summary, summary.sum(axis=0).rename('Total').to_frame().T])

### Write summary report to Excel file

# Open the template, paste the data, autofit the columns and adjust the chart source.
# Save it under a different name.
template = xw.Book(this_dir / "data" / "sales_report_template.xlsx")
sheet = template.sheets['Sheet1']
sheet['B3'].value = summary
sheet['B3'].expand().columns.autofit()
sheet.charts['Chart 1'].set_source_data(sheet['B3'].expand()[:-1, :-1])
template.save(this_dir / 'sales_report_xlwings.xlsx')
