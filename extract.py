import pandas as pd

# Load the Excel file
df = pd.read_excel('jobs.xlsm', sheet_name='Sheet1')

# Convert to JSON
json_data = df.to_json(orient='records')

# Save to a file
with open('jobs.json', 'w') as file:
    file.write(json_data)

print("Data successfully converted to JSON.")
