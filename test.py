import pandas as pd
import numpy as np

# Provided data
reference_list = [24, 29, 21]

data_a = {'x': [24, 29], 'y': ['dog', 'car'], 'z': [None, 'ball']}
data_b = {'x': [29, 21], 'y': ['dog', 'car'], 'z': [None, 'ball']}

a = pd.DataFrame(data_a)
b = pd.DataFrame(data_b)
reference_list = [24, 29, 21]

# Initialize an empty DataFrame 'c' to store matching rows
c = pd.DataFrame(columns=a.columns)
d = pd.DataFrame(columns=a.columns)

# Iterate through reference_list and compare values in all columns
for val in reference_list:
    if val in a['x'].values and val in b['x'].values:
        #matching_row_a = a[a['x'] == val].iloc[0]
        matching_row_a = a[a['x'] == val]
        matching_row_b = b[b['x'] == val]

        matching_row_a = matching_row_a.reset_index(drop=True)
        matching_row_b = matching_row_b.reset_index(drop=True)

        bufferDeltaList = []

        # Check for differences in values and populate delta values accordingly
        for col in a.columns:

            if matching_row_a[col][0] != matching_row_b[col][0]:
                delta_value = f"in df a: {matching_row_a[col][0]} -> in df b: {matching_row_b[col][0]}"

                #c.at[matching_row_a.name, f'delta_{col}'] = delta_value
                d = pd.concat ([c,matching_row_b])

                bufferDeltaList.append(delta_value)

        #d["delta"] = " ".join(str(bufferDeltaList) for x in bufferDeltaList)
        d["delta"] = "".join(str(bufferDeltaList))

# Print the resulting DataFrame 'c'
print("\nDataframe 'd':")
print(d)
