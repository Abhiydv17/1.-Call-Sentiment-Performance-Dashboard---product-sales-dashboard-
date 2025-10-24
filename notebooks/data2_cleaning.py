import pandas as pd

# Step 1: Load the Excel file
df = pd.read_excel(r'c:\Users\HP\Desktop\sales-analysis-capestone\data\Product-Sales-Region.xlsx')  # Adjust sheet_name if needed

# Step 2: Preview structure
print("Initial shape:", df.shape)
print("Columns:", df.columns.tolist())
print(df.head())

# Step 3: Clean column names
df.columns = df.columns.str.strip().str.replace(' ', '_').str.lower()

# Step 4: Drop empty or irrelevant columns
df.dropna(axis=1, how='all', inplace=True)  # Drop columns that are completely empty

# Step 5: Convert data types
if 'date' in df.columns:
    df['date'] = pd.to_datetime(df['date'], errors='coerce')

numeric_cols = ['cost', 'unit_price', 'units_sold']
for col in numeric_cols:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='coerce')

# Step 6: Recalculate revenue and profit if possible
if all(col in df.columns for col in ['units_sold', 'unit_price']):
    df['revenue'] = df['quantity'] * df['unitprice']
    df['profit'] = df['revenue'] * 0.2
    df['quantity'] = pd.to_numeric(df['quantity'], errors='coerce')
df['unitprice'] = pd.to_numeric(df['unitprice'], errors='coerce')
if all(col in df.columns for col in ['revenue', 'cost']):
    df['profit'] = df['revenue'] - df['cost']
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]

# Step 7: Extract month and year for trend analysis
if 'date' in df.columns:
    df['month'] = df['date'].dt.month_name()
    df['year'] = df['date'].dt.year

# Step 8: Drop rows with missing critical values
df.dropna(subset=['date', 'product',], inplace=True)

# Step 9: Save cleaned version (optional)
df.to_csv('cleaned_product_sales.csv', index=False)
print("Cleaned data saved as 'cleaned_product_sales.csv'")