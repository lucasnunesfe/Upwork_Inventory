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

                        df_buffer2 = pd.concat([df_buffer1, matching_row_scan])

                        bufferDeltaList.append(delta_value)
                        #df_buffer2["Status"] = "C"
                        control_change = 1

                if control_change == 1:
                    df_buffer2["Status"] = "C"
                else:
                    continue

                df_buffer2["delta"] = "".join(str(bufferDeltaList))

            if val in reference_df['Barcode_Number'].values and val not in scan_df['Barcode_Number'].values:
                matching_row_reference = reference_df[reference_df['Barcode_Number'] == val]
                df_buffer2 = pd.concat([df_buffer1, matching_row_reference])
                df_buffer2["Status"] = "M"

            if val not in reference_df['Barcode_Number'].values and val in scan_df['Barcode_Number'].values:
                matching_row_scan = scan_df[scan_df['Barcode_Number'] == val]
                df_buffer2 = pd.concat([df_buffer1, matching_row_scan])
                df_buffer2["Status"] = "N"

            df_final = pd.concat([df_final, df_buffer2])

        df_final = df_final.reset_index(drop=True)

        # Save the result to an export file
        export_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        df_final.to_excel(export_file_path, index=False)
        print("\nReconciliation completed. Export file saved at:", export_file_path)

        # Close the GUI window
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = InventoryReconciliationApp(root)
    root.mainloop()
