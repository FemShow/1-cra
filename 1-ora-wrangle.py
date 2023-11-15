import pandas as pd

# Define the file path and worksheet name
file_path = '/Users/femisokoya/Documents/GitHub/CRA/Companies_Register_Activities_2021-22.xls'
worksheet_name = 'Table_B3'

# Read the Excel file, setting row 5 as the header
df = pd.read_excel(file_path, sheet_name=worksheet_name)

# Unmerge the merged cells in row 4
for col in df.columns:
    df.at[4, col] = None

# Drop all the rows below row 14
df = df.iloc[:15]

# List of column indices to process
column_indices = [1, 5, 9, 13, 17, 21, 25, 29, 33, 37]

# Define rows to update based on your requirements
rows_to_update = [5, 8, 9, 10, 11, 12, 13]

# Iterate over each column index
for col_index in column_indices:
    # Set the header in row 3 of the current column to 'Year'
    df.iat[3, col_index] = 'Year'
    
    # Loop through the specified rows and set the values in the current column
    for row in rows_to_update:
        df.iat[row, col_index] = df.iat[2, col_index + 1]

# Make row 5 the column header and delete the top four rows
df = df.iloc[3:].reset_index(drop=True)
df.columns = df.iloc[0]
df = df.iloc[1:].reset_index(drop=True)

# Remove leading and trailing spaces
df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

# Remove blank rows
df = df.dropna()

# Replace values in the 'Corporate body type' column
df['Corporate body type'] = df['Corporate body type'].replace({
    'Assurance Companies2': 'Assurance Companies',
    'Incorporated by Royal Charter4': 'Incorporated by Royal Charter',
    'Special Acts of Parliament5': 'Special Acts of Parliament',
    'Newspaper and Libel Act 18816': 'Newspaper and Libel Act 1881',
    'European Economic Interest Groupings, Principal establishment in UK7, 8, 11': 'European Economic Interest Groupings, Principal establishment in UK',
    'European Public Limited Liability Companies (Societas Europaea)9, 10, 11': 'European Public Limited Liability Companies (Societas Europaea)'
})

# Ensure that column names are unique by adding suffixes
df.columns = [
    'Corporate body type', 'Year_1', 'New_1', 'Closed_1', 'On the register_1',
    'Year_2', 'New_2', 'Closed_2', 'On the register_2',
    'Year_3', 'New_3', 'Closed_3', 'On the register_3',
    'Year_4', 'New_4', 'Closed_4', 'On the register_4',
    'Year_5', 'New_5', 'Closed_5', 'On the register_5',
    'Year_6', 'New_6', 'Closed_6', 'On the register_6',
    'Year_7', 'New_7', 'Closed_7', 'On the register_7',
    'Year_8', 'New_8', 'Closed_8', 'On the register_8',
    'Year_9', 'New_9', 'Closed_9', 'On the register_9',
    'Year_10', 'New_10', 'Closed_10', 'On the register_10'
]

# Function to process columns
def process_columns(df):
    for column in df.columns:
        # Convert column to string
        df[column] = df[column].astype(str)
        
        # Check for value '-' (hyphen) or NaN
        hyphen_values = (df[column] == ' -') | (df[column] == ' - ') | (df[column] == '- ') | (df[column] == '-') | (df[column].isnull())

        if hyphen_values.any():
            # Create a new obsStatus column named as the current column name '-obsStatus'
            obs_status_column_name = f'{column}-obsStatus'

            # Insert a new obsStatus column next to the current column
            df.insert(df.columns.get_loc(column) + 1, obs_status_column_name, '')

            # Replace hyphen values in the original column with a blank value
            df.loc[hyphen_values, column] = ''

            # Insert 'x' in the new obsStatus column on the same rows
            df.loc[hyphen_values, obs_status_column_name] = 'x'

# Function to convert year formats using a dictionary
def convert_year_format(df):
    # Identify columns with 'Year' prefix
    year_columns = [col for col in df.columns if col.startswith('Year')]

    year_conversion_dict = {
        '12-13': '12-2013', '13-14': '13-2014', '14-15': '14-2015',
        '15-16': '15-2016', '16-17': '16-2017', '17-18': '17-2018',
        '18-19': '18-2019', '19-20': '19-2020', '20-21': '20-2021', '21-22': '21-2022'
    }

    for column in year_columns:
        # Replace the two-digit year format
        df[column] = df[column].replace(year_conversion_dict, regex=True)

# Process columns
process_columns(df)

# Convert year formats
convert_year_format(df)

# Concatenate columns vertically
result1 = pd.concat([df['Corporate body type']] * 10, ignore_index=True)
result2 = pd.concat([df[f'Year_{i+1}'].astype(str) for i in range(10)], ignore_index=True)
result3 = pd.concat([df[f'New_{i+1}'].astype(str) for i in range(10)], ignore_index=True)
result4 = pd.concat([df[f'Closed_{i+1}'].astype(str) for i in range(10)], ignore_index=True)
result5 = pd.concat([df[f'On the register_{i+1}'].astype(str) for i in range(10)], ignore_index=True)

# Combine the results into a new DataFrame
result_df = pd.DataFrame({
    'Corporate body type': result1,
    'Year': result2,
    'New': result3,
    'Closed': result4,
    'On_reg': result5
})

# Check for value 'R' suffix and handle blank values on-the-fly
for column in ['New', 'Closed', 'On_reg']:
    r_suffix_values = result_df[column].str.endswith('R')
    blank_values = (result_df[column] == '') | (result_df[column].isnull())

    if r_suffix_values.any():
        # Create a new column named as the current column name suffixed with -revised
        revised_column_name = f'{column}-revised'

        # Insert a new column next to the current column
        result_df.insert(result_df.columns.get_loc(column) + 1, revised_column_name, '')

        # Replace values in the original column with the values without the 'R' suffix
        result_df.loc[r_suffix_values, revised_column_name] = 'R'
        result_df.loc[r_suffix_values, column] = result_df.loc[r_suffix_values, column].str.rstrip('R')

    if blank_values.any():
        # Create a new obsStatus column named as the current column name '-obsStatus'
        obs_status_column_name = f'{column}-obsStatus'

        # Insert a new obsStatus column next to the current column
        result_df.insert(result_df.columns.get_loc(column) + 1, obs_status_column_name, '')

        # Replace blank values in the original column with a blank value
        result_df.loc[blank_values, column] = ''

        # Insert 'x' in the new obsStatus column on the same rows
        result_df.loc[blank_values, obs_status_column_name] = 'x'

# Print the processed DataFrame
print(result_df)
# Save the result in a CSV file
output_file_result = '1-cra-resulta-final.csv'
result_df.to_csv(output_file_result, index=False)

print(f'Data has been processed and saved to {output_file_result}')