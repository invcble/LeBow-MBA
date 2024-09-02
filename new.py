import pandas as pd
import numpy as np

file_path = 'dummy_data_value.xlsx'
df = pd.read_excel(file_path, header=0)

data_dict_path = 'data-dictionary.xlsx'
data_dict_df = pd.read_excel(data_dict_path, header=0)

df = df[df['Finished'] == 1].reset_index(drop=True)

# Function to identify and merge duplicate columns
def merge_columns_with_suffix(df):
    grouped_columns = {}
    
    for col in df.columns:
        base_name = col.split('.')[0]
        
        if base_name in grouped_columns:
            grouped_columns[base_name].append(col)
        else:
            grouped_columns[base_name] = [col]
    
    for base_name, cols in grouped_columns.items():
        if len(cols) > 1:
            df[f'{base_name}_avg'] = df[cols].mean(axis=1, skipna=True)
            df.drop(cols, axis=1, inplace=True)
            df.rename(columns={f'{base_name}_avg': base_name}, inplace=True)
    
    return df

df = merge_columns_with_suffix(df)

recode_columns = data_dict_df[data_dict_df['Recode_Ind'] == 1]['New_Q'].tolist()
# print(recode_columns)

# Subtract 8 from the identified columns to recode
for col in recode_columns:
    if col in df.columns:
        df[col] = 8 - df[col]


categories = data_dict_df.groupby('Category')['New_Q'].apply(list).to_dict()
# print(categories)

network_categories = {key: categories.pop(key) for key in ['N1', 'N2', 'N3', 'N4', 'N5']}
print(network_categories)

# Average the values of columns grouped by category
for category, columns in categories.items():
    available_columns = [col for col in columns if col in df.columns]
    if available_columns:
        try:
            df[category] = df[available_columns].mean(axis=1, skipna=True)
        except:
            print("Dropped columns: ", available_columns)
        df.drop(columns=available_columns, inplace=True)

df.drop(['StartDate', 'EndDate', 'Status', 'IPAddress', 'Progress',
       'Duration (in seconds)', 'Finished', 'RecordedDate', 'ResponseId',
       'RecipientLastName', 'RecipientFirstName', 'RecipientEmail',
       'ExternalReference', 'LocationLatitude', 'LocationLongitude',
       'DistributionChannel', 'UserLanguage'], axis=1, inplace=True)

df = df.dropna(subset=['Q1', 'Q2'])

df['Name'] = (df['Q1'].str.strip() + ' ' + df['Q2'].str.strip()).str.title()
df.drop(['Q1', 'Q2'], axis=1, inplace=True)

df.rename(columns={'Q34': 'Email'}, inplace=True)
df.drop_duplicates(subset=['Email'], keep='last', inplace=True)

mba_program_mapping = {
    1: 'Online MBA',
    2: 'Malvern MBA (Vanguard)',
    3: 'University City MBA',
    4: 'Full-Time MBA'
}

academic_years_mapping = {
    1: '2024-2025',
    2: '2025-2026',
    3: '2026-2027',
    4: '2027-2028',
    5: '2028-2029'
}

academic_term_mapping = {
    1: 'Fall Quarter',
    2: 'Winter Quarter',
    3: 'Spring Quarter'
}

gender_mapping = {
    1: 'Male',
    2: 'Female',
    3: 'Non-binary',
    4: 'Other, or do not wish to share'
}


df.rename(columns={'Q3': 'MBA program', 'Q39': 'Academic years', 'Q40': 'Academic term', 'Q5': 'Gender', 'Q4': 'DOB', 'Q7': 'Profession', 'Q8': 'Work Ex. years'}, inplace=True)
df['MBA program'] = df['MBA program'].map(mba_program_mapping)
df['Academic years'] = df['Academic years'].map(academic_years_mapping)
df['Academic term'] = df['Academic term'].map(academic_term_mapping)
df['Gender'] = df['Gender'].map(gender_mapping)





print(df)
print(df.columns)


# Extract Email and Q23 related data
mba_ntwk = df[['Email'] + [col for sublist in network_categories.values() for col in sublist if col in df.columns]]

# Function to calculate means for a given dimension
def calculate_network_means(df, columns, divisor_value):
    df_filtered = df[columns].copy()
    net_size = df_filtered.notna().sum(axis=1)
    category_count = (df_filtered == divisor_value).sum(axis=1)
    # Avoid division by zero
    return category_count / net_size

# Combine results into a DataFrame
net_breadth = mba_ntwk[['Email']].copy()

# Corrected calculation for individual Cross-Functional (CF), External Contacts (EC), Higher Levels (HLC) averages
net_breadth['Cross_Func'] = calculate_network_means(mba_ntwk, network_categories['N2'], 2)
net_breadth['Ext'] = calculate_network_means(mba_ntwk, network_categories['N3'], 2)
net_breadth['High_Lvl'] = calculate_network_means(mba_ntwk, network_categories['N4'], 3)

# Calculate overall means for CF, EC, and HLC
cf_mean = net_breadth['Cross_Func'].mean(skipna=True)
ec_mean = net_breadth['Ext'].mean(skipna=True)
hlc_mean = net_breadth['High_Lvl'].mean(skipna=True)

# Add average row
average_row = pd.Series(['Average', cf_mean, ec_mean, hlc_mean], index=net_breadth.columns)
net_breadth = pd.concat([net_breadth, average_row.to_frame().T], ignore_index=True)

# Display the final wide DataFrame before melting, if needed
print(net_breadth)


# TODO
# inspect index, regex








# Replace values of 2 with 1 specifically in columns corresponding to N1
n1_columns = network_categories['N1']
mba_ntwk[n1_columns] = mba_ntwk[n1_columns].replace(2, 1)

# Create the subset DataFrame mba_ntwkSS_St_Avg directly using Email and N1 columns
mba_ntwkSS_St_Avg = mba_ntwk[['Email'] + n1_columns]

# Calculate network size (non-missing values count)
mba_ntwkSS_St_Avg['Total_Size'] = mba_ntwkSS_St_Avg.iloc[:, 1:].notna().sum(axis=1)

# Calculate weak and strong ties
mba_ntwkSS_St_Avg['Weak'] = mba_ntwkSS_St_Avg.iloc[:, 1:].apply(lambda row: (row == 1).sum(), axis=1)
mba_ntwkSS_St_Avg['Strong'] = mba_ntwkSS_St_Avg.iloc[:, 1:].apply(lambda row: (row == 3).sum(), axis=1)

# Calculate the average total size
average_total_size = mba_ntwkSS_St_Avg['Total_Size'].mean(skipna=True)
print(f"Average Total Size: {average_total_size}")



# Display the final DataFrame
print(mba_ntwkSS_St_Avg)