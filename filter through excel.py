import pandas as pd
import os

# user inputted variables
file_path = r""
save_path = r""
acceptable_terms = []
column_num_to_be_sorted_on = 0

# program variables
df = pd.read_excel(file_path, names=["Word", "Frequency"])
num_of_rows = df.shape[0]
new_df = pd.DataFrame()

# iterate through rows
for row_number in range(num_of_rows):
    data = df.iat[row_number, column_num_to_be_sorted_on].lower()
    # add if word is acceptable
    for term in acceptable_terms:
        if term in data:
            row_data = df.iloc[row_number]
            new_df = new_df.append(row_data)

os.chdir(save_path)
new_df.to_excel("Filtered Data.xlsx")
exit()
