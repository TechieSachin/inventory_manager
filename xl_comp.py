import pandas as pd


input_agile_path = "input/excel/agile/"
output_excel_path = "output/excel/"
csvname = "output.csv"
xlsxname = "output.xlsx"

# Load the data from the first Excel file
df1 = pd.read_excel(input_agile_path+"Invoice_ SO-000078.xlsx")
#print(df1)
# Load the data from the second Excel file
df2 = pd.read_excel(input_agile_path+"Invoice_ SO-000139.xlsx")

# Merge the two dataframes on the item names (assuming the item names are in a column called 'Item')
merged_data = df1.merge(df2, on='Product', suffixes=('_file1', '_file2'))

# Compare the item prices and create a new column for the price difference
merged_data['Price_Difference'] = merged_data['Unit Price_file1'] - merged_data['Unit Price_file2']

# You can now view or export the results
print(merged_data)

# Export the results to a new Excel file
merged_data.to_excel('price_comparison_result.xlsx', index=False)

# Read the existing XLSX file
df = pd.read_excel('price_comparison_result.xlsx')

# Keep only the desired columns
columns_to_keep = ['Product', 'Price_Difference', 'Unit Price_file1', 'Unit Price_file2']
df = df[columns_to_keep]

# Write the updated DataFrame back to the XLSX file
df.to_excel('price_comparison_result.xlsx', index=False, engine='openpyxl')

print("Columns have been filtered and saved back to 'price_comparison_result.xlsx'.")




# Load the Excel file
file_path = input_agile_path+"Invoice_ SO-000106.xlsx"
df = pd.read_excel(file_path)

# Function to convert a dollar amount string to a number
def dollar_string_to_number(dollar_string):
    if isinstance(dollar_string, str):
        try:
            # Remove dollar sign and commas, then convert to float
            return float(dollar_string.replace('$', '').replace(',', ''))
        except ValueError:
            return dollar_string  # Return the original value if conversion fails
    return dollar_string  # Return the original value if it's not a string


# Apply the function to the 'Amount' column
df['Unit Price'] = df['Unit Price'].apply(dollar_string_to_number)

# Save the modified DataFrame back to the same Excel file
df.to_excel(file_path, index=False, engine='openpyxl')

print("Column 'Amount' in the Excel file has been updated with numeric values.")
