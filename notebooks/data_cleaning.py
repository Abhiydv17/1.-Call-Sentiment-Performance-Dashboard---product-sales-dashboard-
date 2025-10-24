import pandas as pd

# Step 1: Load the Excel file without header
df_raw = pd.read_excel("C:/Users/HP/Desktop/sales-analysis-capestone/data/sentiment_data.xlsx", header=None)

# Step 2: Manually set the correct column names from row 6 (index 6)
df_raw.columns = df_raw.iloc[6]

# Step 3: Drop the first 7 rows (junk rows + header row)
df = df_raw.iloc[7:].copy()

# Step 4: Reset index
df.reset_index(drop=True, inplace=True)

# Step 5: Show cleaned data
print(df.head())
print("\nColumns:", df.columns.tolist())

# Step 6: Rename columns manually
df.columns = [
    'Index', 'ID', 'Customer Name', 'Sentiment', 'CSAT Score',
    'Call Timestamp', 'Reason', 'City', 'State', 'Channel',
    'Response Time', 'Call Duration (Minutes)', 'Call Center'
]

# Confirm structure
print("\nCleaned Columns:", df.columns.tolist())
print("\nSample Data:\n", df.head())

# Step 7: Check for missing values
missing = df.isnull().sum()
print("\nMissing values per column:\n", missing)

# Step 8: Save cleaned data to CSV
df.to_csv("C:/Users/HP/Desktop/sales-analysis-capestone/data/sentiment_cleaned.csv", index=False)
print("\nCleaned data saved to sentiment_cleaned.csv")