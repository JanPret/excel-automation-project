import pandas as pd

# Step 1: Read CSV
df = pd.read_csv("data/sales_data.csv")

# Step 2: Clean Data
df.drop_duplicates(inplace=True)
df.fillna(0, inplace=True)

# Step 3: Create Summary Tables
sales_by_product = df.groupby('Product')['Amount'].sum().reset_index()
sales_by_month = df.groupby('Month')['Amount'].sum().reset_index()

# Step 4: Write to Excel
with pd.ExcelWriter("output/sales_report.xlsx", engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name='Raw Data', index=False)
    sales_by_product.to_excel(writer, sheet_name='Sales by Product', index=False)
    sales_by_month.to_excel(writer, sheet_name='Sales by Month', index=False)

    workbook = writer.book
    worksheet = writer.sheets['Sales by Product']
    chart = workbook.add_chart({'type': 'column'})
    chart.add_series({
        'name': 'Sales by Product',
        'categories': ['Sales by Product', 1, 0, len(sales_by_product), 0],
        'values': ['Sales by Product', 1, 1, len(sales_by_product), 1],
    })
    worksheet.insert_chart('D2', chart)

print("Excel report generated successfully!")
