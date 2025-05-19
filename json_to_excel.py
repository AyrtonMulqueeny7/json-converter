import json
import pandas as pd
import sqlite3

# Load JSON data
with open('out.json', 'r') as f:
    data = json.load(f)

# Convert to DataFrame
df = pd.DataFrame(data)

# Flatten nested data by converting lists to strings
for col in df.columns:
    df[col] = df[col].apply(lambda x: json.dumps(x) if isinstance(x, (dict, list)) else x)

print("Data columns:", df.columns.tolist())
print("Number of records:", len(df))

# Connect to SQLite and write data
conn = sqlite3.connect(':memory:')
df.to_sql('your_table', conn, index=False)

# Run query
query = "SELECT * FROM your_table"
result = pd.read_sql_query(query, conn)

# Export to Excel
result.to_excel('output.xlsx', index=False)
print("Excel file created successfully!")
