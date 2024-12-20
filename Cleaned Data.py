import pandas as pd

# Load the Excel file
file_path = "employee_burnout_analysis.xlsx"
df = pd.read_excel(file_path, sheet_name="in")

# Handle missing values by filling numeric columns with their median
numeric_columns = ['Resource Allocation', 'Mental Fatigue Score', 'Burn Rate']
for col in numeric_columns:
    df[col].fillna(df[col].median(), inplace=True)

# Convert "Date of Joining" to datetime format
df['Date of Joining'] = pd.to_datetime(df['Date of Joining'])

# Standardize categorical columns
df['Gender'] = df['Gender'].str.capitalize()
df['WFH Setup Available'] = df['WFH Setup Available'].str.capitalize()

# Check for duplicates and remove them
df.drop_duplicates(inplace=True)

# Save the cleaned data back to a new Excel file
output_path = "cleaned_employee_data.xlsx"
df.to_excel(output_path, index=False)

print(f"Data cleaned and saved to {output_path}")
