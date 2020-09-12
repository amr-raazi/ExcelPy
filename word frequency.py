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
    master_df = master_df.append(pd.read_excel(file, header=None, skiprows=25), ignore_index=True)

# define new dataframe for frequency data
frequency_data = master_df.iloc[:, frequency_data_column_number].dropna().to_frame()

# set header for frequency data
frequency_data.columns = ["data"]

# get frequency data from master dataframe
frequency = frequency_data.data.str.split(expand=True).stack().value_counts().to_frame()

# save to excel
frequency.to_excel("Word Frequency.xlsx")
