import tkinter as tk
from tkinter import filedialog
import pandas as pd
import numpy as np
import openpyxl

class InventoryReconciliationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Inventory Reconciliation Tool")

        self.reference_file_path = ""
        self.scan_file_path = ""

        # UI Elements
        self.label_reference = tk.Label(root, text="Reference File:")
        self.label_reference.grid(row=0, column=0)

        self.label_scan = tk.Label(root, text="Scan File:")
        self.label_scan.grid(row=1, column=0)

        self.btn_reference = tk.Button(root, text="Select Reference File", command=self.load_reference_file)
        self.btn_reference.grid(row=0, column=1)

        self.btn_scan = tk.Button(root, text="Select Scan File", command=self.load_scan_file)
        self.btn_scan.grid(row=1, column=1)

        self.btn_reconcile = tk.Button(root, text="Reconcile", command=self.reconcile_files)
        self.btn_reconcile.grid(row=2, column=0, columnspan=2)

    def load_reference_file(self):
        self.reference_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])

    def load_scan_file(self):
        self.scan_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])

    def reconcile_files(self):
        if not self.reference_file_path or not self.scan_file_path:
            return

        reference_df = pd.read_excel(self.reference_file_path)
        scan_df = pd.read_excel(self.scan_file_path)

        # Replace spaces in column names with underscores
        reference_df.columns = [col.replace(' ', '_') for col in reference_df.columns]
        scan_df.columns = [col.replace(' ', '_') for col in scan_df.columns]



        # Check and convert to the same variable type if needed
        '''for col in ['COST', 'Funding_Code']:
            if col in reference_df.columns and col in scan_df.columns:
                try:
                    # Skip conversion if the value is missing
                    if pd.notna(scan_df[col]).all():
                        scan_df[col] = scan_df[col].astype('float64')
                except Exception as e:
                    print(f"Error converting column '{col}' to the same type. Details: {str(e)}")'''

        # Merge dataframes based on Barcode Number column
        merged_df = pd.merge(reference_df, scan_df, left_on='Barcode_Number', right_on='Barcode_Number', how='outer',
                             suffixes=('_reference', '_scan'))

        ###included
        # Get column 'x' from dataframes 'a' and 'b' and turn into a list
        column_BarcodeNumber_reference = list(reference_df['Barcode_Number'])
        column_BarcodeNumber_scan = list(scan_df['Barcode_Number'])

        # Combine the lists and remove duplicates
        unique_values = list(set(column_BarcodeNumber_reference + column_BarcodeNumber_scan))

        df_buffer1 = pd.DataFrame(columns=reference_df.columns)
        df_buffer2 = pd.DataFrame(columns=reference_df.columns)
        df_final = pd.DataFrame(columns=reference_df.columns)
        df_final["delta"] = None
        df_final["Status"] = None

        for val in unique_values:
            if val in reference_df['Barcode_Number'].values and val in scan_df['Barcode_Number'].values:
                matching_row_reference = reference_df[reference_df['Barcode_Number'] == val]
                matching_row_scan = scan_df[scan_df['Barcode_Number'] == val]

                matching_row_reference = matching_row_reference.reset_index(drop=True)
                matching_row_scan = matching_row_scan.reset_index(drop=True)

                bufferDeltaList = []

                control_change = 0

                # Check for differences in values and populate delta values accordingly
                for col in reference_df.columns:

                    if matching_row_reference[col][0] != matching_row_scan[col][0]:
                        delta_value = f"in df a: {matching_row_reference[col][0]} -> in df b: {matching_row_scan[col][0]}"

                        # c.at[matching_row_a.name, f'delta_{col}'] = delta_value
                        df_buffer2 = pd.concat([df_buffer1, matching_row_scan])

                        bufferDeltaList.append(delta_value)
                        #df_buffer2["Status"] = "C"
                        control_change = 1

                    #elif matching_row_reference[col][0] == matching_row_scan[col][0]:
                        #df_buffer2["Status"] = "F"

                if control_change == 1:
                    df_buffer2["Status"] = "C"
                else:
                    continue

                # d["delta"] = " ".join(str(bufferDeltaList) for x in bufferDeltaList)
                df_buffer2["delta"] = "".join(str(bufferDeltaList))

            #df_buffer2 = pd.DataFrame(columns=reference_df.columns)

            if val in reference_df['Barcode_Number'].values and val not in scan_df['Barcode_Number'].values:
                matching_row_reference = reference_df[reference_df['Barcode_Number'] == val]
                df_buffer2 = pd.concat([df_buffer1, matching_row_reference])
                df_buffer2["Status"] = "M"

            if val not in reference_df['Barcode_Number'].values and val in scan_df['Barcode_Number'].values:
                matching_row_scan = scan_df[scan_df['Barcode_Number'] == val]
                df_buffer2 = pd.concat([df_buffer1, matching_row_scan])
                df_buffer2["Status"] = "N"


                #df_buffer2.at[df.index[df['BoolCol']].tolist(), 'Status'] = 'M'  # Missing
            df_final = pd.concat([df_final, df_buffer2])

        df_final = df_final.reset_index(drop=True)

        ##nao está em uso
        # Create a new column for status
        merged_df['Status'] = ''

        # Identify Missing and New items
        '''for index, row in merged_df.iterrows():
            barcode_number = row['Barcode_Number']

            if barcode_number in reference_df['Barcode_Number'].values and barcode_number not in scan_df[
                'Barcode_Number'].values:
                merged_df.at[index, 'Status'] = 'M'  # Missing

            if barcode_number not in reference_df['Barcode_Number'].values and barcode_number in scan_df[
                'Barcode_Number'].values:
                merged_df.at[index, 'Status'] = 'N'  # New Item'''

        # Identify Found and Changed items
        for index, row in merged_df.iterrows():
            reference_column = 'Barcode_Number_reference'
            scan_column = 'Barcode_Number_scan'

            if reference_column not in merged_df.columns or scan_column not in merged_df.columns:
                continue

            barcode_number = row['Barcode_Number']

            if pd.notna(row[reference_column]) and pd.notna(row[scan_column]):
                # Use numpy for efficient element-wise comparison of rows
                reference_row = row.filter(like=reference_column).to_numpy()
                scan_row = row.filter(like=scan_column).to_numpy()

                # Check for equality, considering NaN values
                if np.array_equal(reference_row, scan_row):
                    merged_df.at[index, 'Status'] = 'F'  # Found
                else:
                    merged_df.at[index, 'Status'] = 'C'  # Changed

        # Print the entire merged DataFrame
        pd.set_option('display.max_rows', None)
        print("\nMerged DataFrame:")
        print(merged_df)
        pd.reset_option('display.max_rows')  # Reset the option to its default value

        # Replace spaces in column names with underscores and add suffixes
        reference_df.columns = [col.replace(' ', '_') + '_reference' if col != 'Barcode_Number' else col for col in
                                reference_df.columns]
        scan_df.columns = [col.replace(' ', '_') + '_scan' if col != 'Barcode_Number' else col for col in
                           scan_df.columns]

        # Extract original values for Missing and New items
        missing_items_barcodes = merged_df.loc[merged_df['Status'] == 'M', 'Barcode_Number']
        new_items_barcodes = merged_df.loc[merged_df['Status'] == 'N', 'Barcode_Number']

        missing_items = reference_df[reference_df['Barcode_Number'].isin(missing_items_barcodes)].copy()
        missing_items['Status'] = 'M'  # Set 'Status' for Missing items

        new_items = scan_df[scan_df['Barcode_Number'].isin(new_items_barcodes)].copy()
        new_items['Status'] = 'N'  # Set 'Status' for New items

        # Concatenate the extracted values
        export_df = pd.concat([missing_items, new_items], axis=0, ignore_index=True)

        # Reorder columns to have 'Status' as the first column
        export_df = export_df[['Status'] + [col for col in export_df.columns if col != 'Status']]

        # Save the result to an export file
        export_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        export_df.to_excel(export_file_path, index=False)
        print("\nReconciliation completed. Export file saved at:", export_file_path)

        # Close the GUI window
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = InventoryReconciliationApp(root)
    root.mainloop()
