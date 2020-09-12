import pandas
import os

# user inputted variables
folder_path = r"excel_files"
data_row_number = 1

# program variable
file_list = os.listdir(folder_path)

# cd into directory
os.chdir(folder_path)

# create empty dataframe
master_df = pandas.DataFrame()

# go through files and append data into master dataframe
for file in file_list:
    if file.endswith('.xlsx'):
        master_df = master_df.append(pandas.read_excel(file, header=None, skiprows=1), ignore_index=True, )

# set new header for the row whose word frequency is to be printed
try:
    new_header = list(master_df.columns.values)
    new_header[data_row_number] = "data"
    master_df.columns = new_header
except IndexError:
    print("Row specified is not in the given columns or no file given")
    exit()

# get frequency data from master dataframe
frequency = master_df.data.str.split(expand=True).stack().value_counts()

# save to excel
pandas.frequency.to_excel("word frequency.xlsx")
