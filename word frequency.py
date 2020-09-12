import pandas as pd
import os

# user inputted variables
folder_path = r"excel_files"
frequency_data_column_number = 3

# program variable
file_list = os.listdir(folder_path)

# cd into directory
os.chdir(folder_path)

# create empty dataframe
master_df = pd.DataFrame()

# go through files and append data into master dataframe
for file in file_list:
    master_df = master_df.append(pd.read_excel(file, header=None, skiprows=25), ignore_index=True, )

# set new header for the row whose word frequency is to be printed in the master dataframe
try:
    new_header = list(master_df.columns.values)
    new_header[data_row_number] = "data"
    master_df.columns = new_header
except IndexError:
    print("Row specified is not in the given columns or no file given")
    exit()

# define new dataframe for frequency data
frequency_data = master_df["data"].dropna().to_frame()

# get frequency data from master dataframe
frequency = frequency_data.data.str.split(expand=True).stack().value_counts().to_frame()

# save to excel
frequency.to_excel("Word Frequency.xlsx")
