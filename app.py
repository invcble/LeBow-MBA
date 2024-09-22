import customtkinter as ctk
import pandas as pd
import os
import warnings
from tkinter import messagebox

pd.set_option('future.no_silent_downcasting', True)
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
warnings.filterwarnings("ignore", message=".*Tight layout not applied.*")

from gen_report import save_reports


class MBAReportApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Window configuration
        self.title("MBA Report Generator")
        self.geometry("600x450")

        # File path inputs
        self.file_path_label = ctk.CTkLabel(self, text="Survey Data File (.xlsx) Path:")
        self.file_path_label.pack(pady=10)
        self.file_path_entry = ctk.CTkEntry(self, width=500)
        self.file_path_entry.pack()

        self.data_dict_path_label = ctk.CTkLabel(self, text="Database Map File (.xlsx) Path:")
        self.data_dict_path_label.pack(pady=10)
        self.data_dict_entry = ctk.CTkEntry(self, width=500)
        self.data_dict_entry.pack()

        self.template_path_label = ctk.CTkLabel(self, text="Template File (.pdf) Path:")
        self.template_path_label.pack(pady=10)
        self.template_entry = ctk.CTkEntry(self, width=500)
        self.template_entry.pack()

        # Term entry
        self.term_label = ctk.CTkLabel(self, text="Enter New Term (e.g., Fall 21, Spring 22, Winter 23):")
        self.term_label.pack(pady=10)
        self.term_entry = ctk.CTkEntry(self, width=500)
        self.term_entry.pack()

        # Submit button
        self.check_box = ctk.CTkCheckBox(self, text="Late Submission Mode (Doesn't update Term & MBA Average)")
        self.check_box.pack(pady=20)

        # Submit button
        self.submit_button = ctk.CTkButton(self, text="Generate Report", command=self.generate_report)
        self.submit_button.pack(pady=20)

    def merge_columns_with_suffix(self, df):
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

    def calculate_network_means(self, df, columns, divisor_value):
        df_filtered = df[columns].copy()
        net_size = df_filtered.notna().sum(axis=1)
        category_count = (df_filtered == divisor_value).sum(axis=1)
        # Avoid division by zero
        return category_count / net_size

    def generate_report(self):
        # Retrieve input values
        file_path = self.file_path_entry.get().replace('"','')
        data_dict_path = self.data_dict_entry.get().replace('"','')
        template_path = self.template_entry.get().replace('"','')
        new_term = self.term_entry.get()

        # Validate inputs
        if not all([file_path, data_dict_path, template_path, new_term]):
            messagebox.showerror("Input Error", "All fields must be filled in.")
            return

        if not all(map(os.path.exists, [file_path, data_dict_path, template_path])):
            messagebox.showerror("File Error", "One or more file paths are invalid.")
            return
        
        if not file_path.endswith('.xlsx'):
            messagebox.showerror("Format Error", "Survey data file must have .xlsx format.")
            return

        # Load DataFrames
        try:
            df = pd.read_excel(file_path, header=0)
            data_dict_df = pd.read_excel(data_dict_path, sheet_name='Map', header=0)
            database_df = pd.read_excel(data_dict_path, sheet_name='Database', header=0)
        except Exception as e:
            messagebox.showerror("Load Error", f"Error loading files: {e}")
            return

        try:
            # Filtering actual responses/ Removing question and tag rows
            df['Progress'] = pd.to_numeric(df['Progress'], errors='coerce')
            df = df[df['Progress'] >= 75].reset_index(drop=True)

            # Function to identify and merge duplicate columns
            df = self.merge_columns_with_suffix(df)

            # Subtract 8 from the identified columns to recode
            recode_columns = data_dict_df[data_dict_df['Recode_Ind'] == 1]['New_Q'].tolist()
            for col in recode_columns:
                if col in df.columns:
                    df[col] = 8 - df[col]

            ################## CATEGORY ##################

            categories = data_dict_df.groupby('Category')['New_Q'].apply(list).to_dict()
            network_categories = {key: categories.pop(key) for key in ['N1', 'N2', 'N3', 'N4', 'N5']}

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
                   'DistributionChannel', 'UserLanguage'], axis=1, inplace=True, errors='ignore')

            # Process Name, Email and other demographics
            df = df.dropna(subset=['Q1', 'Q2'])
            df['Name'] = (df['Q1'].str.strip() + ' ' + df['Q2'].str.strip()).str.title()
            df.drop(['Q1', 'Q2'], axis=1, inplace=True)

            df.rename(columns={'Q34': 'Email', 'Q39': 'Academic years', 'Q40': 'Academic term', 'Q5': 'Gender', 'Q4': 'DOB', 'Q7': 'Profession', 'Q8': 'Work Ex. years'}, inplace=True)
            df.drop_duplicates(subset=['Email'], keep='last', inplace=True)

            mba_program_mapping = {1: 'Online MBA', 2: 'Malvern MBA (Vanguard)', 3: 'University City MBA', 4: 'Full-Time MBA'}
            academic_years_mapping = {1: '2024-2025', 2: '2025-2026', 3: '2026-2027', 4: '2027-2028', 5: '2028-2029'}
            academic_term_mapping = {1: 'Fall Quarter', 2: 'Winter Quarter', 3: 'Spring Quarter'}
            gender_mapping = {1: 'Male', 2: 'Female', 3: 'Non-binary', 4: 'Other, or do not wish to share'}

            df['Academic years'] = df['Academic years'].map(academic_years_mapping)
            df['Academic term'] = df['Academic term'].map(academic_term_mapping)
            df['Gender'] = df['Gender'].map(gender_mapping)
            df['PS'] =  round((df['NA1'] + df['II'] + df['SA'] + df['AS'])/4, 2)
            df['EmpL'] =  round((df['LBE'] + df['PDM'] + df['COACH'] + df['INF'] + df['ShowCon'])/5, 2)
            df['PE'] =  round((df['MEAN'] + df['COMP'] + df['SD'] + df['IMP'])/4, 2)

            ################## NETWORK ##################

            mba_ntwk = df[['Email'] + [col for sublist in network_categories.values() for col in sublist if col in df.columns]].copy()

            # Calculate network means
            mba_ntwk['Cross_Func'] = self.calculate_network_means(mba_ntwk, network_categories['N2'], 2)
            mba_ntwk['Ext'] = self.calculate_network_means(mba_ntwk, network_categories['N3'], 2)
            mba_ntwk['High_Lvl'] = self.calculate_network_means(mba_ntwk, network_categories['N4'], 3)

            # Replace values of 2 with 1 specifically in columns corresponding to N1
            n1_columns = network_categories['N1']
            mba_ntwk[n1_columns] = mba_ntwk[n1_columns].replace(2, 1)

            # Calculate network size (non-missing values count)
            mba_ntwk['Total_Size'] = mba_ntwk.iloc[:, 1:21].notna().sum(axis=1)
            mba_ntwk['Weak'] = mba_ntwk.iloc[:, 1:21].apply(lambda row: (row == 1).sum(), axis=1)
            mba_ntwk['Strong'] = mba_ntwk.iloc[:, 1:21].apply(lambda row: (row == 3).sum(), axis=1)

            mba_ntwk = mba_ntwk[['Email', 'Cross_Func', 'Ext', 'High_Lvl', 'Weak', 'Strong', 'Total_Size']]
            columns_to_drop = [col for sublist in network_categories.values() for col in sublist]
            df = df.drop(columns=columns_to_drop)

            df = pd.merge(df, mba_ntwk, on='Email', how='left').sort_index()
            df = df[['Name'] + [col for col in df.columns if col != 'Name']]


            demographic_columns = ['Name', 'Email', 'Academic years', 'Academic term', 'DOB', 'Gender', 'Profession', 'Work Ex. years']
        
            # Check for Late Mode
            if self.check_box.get():
                print('late mode enabled')

                # No need to calculate Term avg, as we will be loading MBA avg from DB directly in this mode.
            else:
                print('late mode disabled')

                ################## CALCULATE TERM AVERAGE ##################

                numeric_df = df.apply(pd.to_numeric, errors='coerce')
                term_average_dict = numeric_df.drop(columns=demographic_columns, errors='ignore').mean().round(2).to_dict()
                term_average_dict['Weak'] =  round(term_average_dict['Weak'])
                term_average_dict['Strong'] =  round(term_average_dict['Strong'])
                term_average_dict['Total_Size'] =  round(term_average_dict['Total_Size'])
                term_average_dict['PS'] =  round((term_average_dict['NA1'] + term_average_dict['II'] + term_average_dict['SA'] + term_average_dict['AS'])/4, 2)
                term_average_dict['EmpL'] =  round((term_average_dict['LBE'] + term_average_dict['PDM'] + term_average_dict['COACH'] + term_average_dict['INF'] + term_average_dict['ShowCon'])/5, 2)
                term_average_dict['PE'] =  round((term_average_dict['MEAN'] + term_average_dict['COMP'] + term_average_dict['SD'] + term_average_dict['IMP'])/4, 2)

                ################## LOAD AND UPDATE DB ##################

                # Recalculate the 'MBA Average' column, now including new term
                column_titles = database_df.columns.tolist()
                columns_to_remove = ['Sub-Categories', 'Code', 'MBA Average']
                terms = [col for col in column_titles if col not in columns_to_remove]

                if new_term in terms:
                    ans = messagebox.askyesno("Overwrite Term", f"Term '{new_term}' already exists. Do you want to overwrite with new data and update MBA average?")
                    if not ans:
                        return  # Do not proceed
                    else:
                        # Overwrite term
                        database_df[new_term] = database_df['Code'].map(term_average_dict)
                else:
                    # Inserting new term
                    terms.append(new_term)
                    database_df[new_term] = database_df['Code'].map(term_average_dict)

                database_df['MBA Average'] = database_df[terms].mean(axis=1, skipna=True).round(2)

                # Round 'MBA Average' to integers for specific 'Code' values
                codes_to_round = ['Total_Size', 'Strong', 'Weak']
                database_df.loc[database_df['Code'].isin(codes_to_round), 'MBA Average'] = (
                    database_df.loc[database_df['Code'].isin(codes_to_round), 'MBA Average'].round(0).astype(int)
                )

                try:
                    with pd.ExcelWriter(data_dict_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        database_df.to_excel(writer, sheet_name='Database', index=False)
                        print("Database updated.")
                except PermissionError as e:
                    messagebox.showerror("Permission Error", f"{e}\nClose the Database file before running this script.")
                    return

            ################## LOAD MBA AVERAGE FROM UPDATED DB AND SAVE TO DF ##################

            mba_average_dict = {f"{code}_MBAavg": avg for code, avg in database_df.set_index('Code')['MBA Average'].items()}

            for key, value in mba_average_dict.items():
                df[key] = value

            ################## CALCULATE STANDARD DEVIATION ##################

            demographic_columns = ['Name', 'Email', 'Academic years', 'Academic term', 
                               'DOB', 'Gender', 'Profession', 'Work Ex. years']
            strength_columns = ['Strong', 'Weak']
            breadth_columns = ['High_Lvl', 'Ext', 'Cross_Func']

            exclude_cols = demographic_columns + strength_columns + breadth_columns
            unique_columns = [col for col in df.columns if col not in exclude_cols and not col.endswith('_MBAavg')]

            # SD for each unique column using MBA averages
            for col in unique_columns:
                mba_avg_col = f'{col}_MBAavg'

                if mba_avg_col in df.columns:
                    df[f'{col}_SD'] = ((df[col] - df[mba_avg_col]) ** 2).mean() ** 0.5

            ################## SORT MAIN DF ##################

            remaining_columns = sorted([col for col in df.columns if col not in demographic_columns])
            sorted_columns = demographic_columns + remaining_columns

            df = df[sorted_columns]

            with pd.ExcelWriter(f"{new_term}_dataframe.xlsx", engine='openpyxl') as writer:
                df.to_excel(writer, index=False)

            ################## GENERATE REPORTS ##################

            save_reports(new_term, df, template_path)
            print("Reports generated successfully.")
            messagebox.showinfo("Success", "Reports generated successfully.")

        except Exception as e:
            print(e)
            messagebox.showerror("Processing Error", f"Error during processing: {e}")
            return
        
        finally:
            self.focus()

# Run the application
if __name__ == "__main__":
    ctk.set_appearance_mode("Dark")
    app = MBAReportApp()
    app.mainloop()



#TODO
# late submission option
# line by line check
