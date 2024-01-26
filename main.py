import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import numpy as np
import openpyxl

class InventoryReconciliationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Inventory Reconciliation Tool")
        self.root.geometry("220x100")  # Set the initial size of the window

        self.reference_file_path = ""
        self.scan_file_path = ""

        # Add some padding on the top
        self.root.grid_rowconfigure(1, pad=15)

        # UI Elements with left margin
        self.label_reference = tk.Label(root, text="Reference File:")
        self.label_reference.grid(row=1, column=0, sticky="w", padx=(10, 0))  # Add left margin

        self.label_scan = tk.Label(root, text="Scan File:")
        self.label_scan.grid(row=2, column=0, sticky="w", padx=(10, 0))  # Add left margin

        self.btn_reference = tk.Button(root, text="Select Reference File", command=self.load_reference_file)
        self.btn_reference.grid(row=1, column=1, sticky="ew")  # Align button to the right

        self.btn_scan = tk.Button(root, text="Select Scan File", command=self.load_scan_file)
        self.btn_scan.grid(row=2, column=1, sticky="ew")  # Align button to the right

        self.btn_reconcile = tk.Button(root, text="Run Reconciliation", command=self.reconcile_files, bg='black', fg='white')
        self.btn_reconcile.grid(row=3, column=1, columnspan=2, sticky="ew")  # Align button to the right

    def load_reference_file(self):
        self.reference_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])

    def load_scan_file(self):
        self.scan_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])

    def reconcile_files(self):
        if not self.reference_file_path or not self.scan_file_path:
            return

    def load_reference_file(self):
        self.reference_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])

    def load_scan_file(self):
        self.scan_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])

    def reconcile_files(self):
        if not self.reference_file_path or not self.scan_file_path:
            return

        # Read data from reference and scan files
        reference_df = pd.read_excel(self.reference_file_path)
        scan_df = pd.read_excel(self.scan_file_path)

        # Replace spaces in column names with underscores
        reference_df.columns = [col.replace(' ', '_') for col in reference_df.columns]
        scan_df.columns = [col.replace(' ', '_') for col in scan_df.columns]

        # Get column 'x' from dataframes from reference and scan files and turn them into lists
        column_BarcodeNumber_reference = list(reference_df['Barcode_Number'])
        column_BarcodeNumber_scan = list(scan_df['Barcode_Number'])

        # Combine the lists and remove duplicates
        unique_values = list(set(column_BarcodeNumber_reference + column_BarcodeNumber_scan))

        # Initialize DataFrames to store results
        df_buffer1 = pd.DataFrame(columns=reference_df.columns)
        df_buffer2 = pd.DataFrame(columns=reference_df.columns)
        # df_final is the dataframe covering the table to be exported in excel file
        df_final = pd.DataFrame(columns=reference_df.columns)
        df_final["delta"] = None
        df_final["Status"] = None

        # Iterate each existing Barcode Numer (regardless it is from reference or scan files)
        for val in unique_values:

            # If Barcode Number is available in both reference and scan files
            if val in reference_df['Barcode_Number'].values and val in scan_df['Barcode_Number'].values:

                # Matching_row is a single-line dataframe comprising the Barcode Number being evaluated
                matching_row_reference = reference_df[reference_df['Barcode_Number'] == val]
                matching_row_scan = scan_df[scan_df['Barcode_Number'] == val]

                # Resetting index to make sure they are 0-labeled
                matching_row_reference = matching_row_reference.reset_index(drop=True)
                matching_row_scan = matching_row_scan.reset_index(drop=True)

                # Support list to be used in logic below
                bufferDeltaList = []

                # Variable "flag" to monitor if there is any column with different values
                # between reference and scan files
                # 0 means there is no different columns' values
                control_change = 0

                df_buffer2 = pd.concat([df_buffer1, matching_row_scan])

                # Check for differences in values and populate delta values accordingly
                for col in reference_df.columns:
                    # If-condifiton to evalute if there is at least one difference between reference
                    # and scan files
                    if matching_row_reference[col][0] != matching_row_scan[col][0]:
                        if pd.isna(matching_row_reference[col][0]) and pd.isna(matching_row_scan[col][0]):
                            continue
                        else:
                            # Prepare log input to be filled in "Change" column in exported file
                            delta_value = f"ref {col}: {matching_row_reference[col][0]} -> scan {col}: {matching_row_scan[col][0]}"
                            # Stores all tracked differences for that same Barcode Number
                            bufferDeltaList.append(delta_value)
                            # Sets the flag to 1, indicating there is at least 1 column with
                            # Different values
                            control_change = 1

                    # If-condition to evaluate if the entry in reference and scan files are
                    # exactly the same
                    elif matching_row_reference[col][0] == matching_row_scan[col][0]:
                        continue

                # At the end of all columns evaluation for that Barcode Number
                # it is then defined Changed/different values are present
                if control_change == 1:
                    # Flag "C" added to new column from export file called Status
                    df_buffer2["Status"] = "C"
                else:
                    df_buffer2["Status"] = "F"

                # Transform list into str and input "delta" column
                df_buffer2["delta"] = "".join(str(bufferDeltaList))

            # If Barcode Number is missing from scan and present in reference file
            if val in reference_df['Barcode_Number'].values and val not in scan_df['Barcode_Number'].values:
                matching_row_reference = reference_df[reference_df['Barcode_Number'] == val]
                df_buffer2 = pd.concat([df_buffer1, matching_row_reference])
                df_buffer2["Status"] = "M"

            # If Barcode Number is present in reference and missing from scan file
            if val not in reference_df['Barcode_Number'].values and val in scan_df['Barcode_Number'].values:
                matching_row_scan = scan_df[scan_df['Barcode_Number'] == val]
                df_buffer2 = pd.concat([df_buffer1, matching_row_scan])
                df_buffer2["Status"] = "N"

            # Append/concat df_final
            df_final = pd.concat([df_final, df_buffer2])

        # Reset index to make it ok
        df_final = df_final.reset_index(drop=True)

        # Save the result to an export file
        export_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        df_final.to_excel(export_file_path, index=False)
        print("\nReconciliation completed. Export file saved at:", export_file_path)

        # Change the color of the "Run Reconciliation" button to dark gray
        self.btn_reconcile.configure(bg="dark gray")

        # Display message box after reconciliation is completed
        messagebox.showinfo("Reconciliation Completed",
                            "Reconciliation completed. Export file saved at:\n" + export_file_path)

        # Close the GUI window only when the user presses "OK" in the message box
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = InventoryReconciliationApp(root)
    root.mainloop()