import pandas as pd
import numpy as np
import os
from datetime import datetime


# Data frame referred to as "df" in comments
# Enter path to folder
path = "P:\_Common\Sales_Quotes_2020\Jason Smith"
# Create data frame instance to add all pricing (with added intro column) data
all_data = pd.DataFrame()
current_file_intro_data = pd.DataFrame()
current_file_pricing_data = pd.DataFrame()
# Uncomment file counter variables below for testing
# filesAdded = 0
# totalFiles = 0
colNames = ["Quantity", "Sartorius Part Number", "Description", "List Price", "Discount %", "Minimum Sell Price", "Unit Price", "Total Ext Price"]
# Add column names to intro sheet data for
colNames2 = ["A", "B", "C", "D", "E", "F"]
# Grab current time to measure length of process
date_time_obj = datetime.now()
print(date_time_obj)
for root, directories, filenames in os.walk(path):
    for filename in filenames:
        if filename.endswith(".xlsx"):
            # Uncomment file counter below for testing total files read
            # totalFiles += 1
            try:
                # Designate which sheets, start col/row, col names, # of cols/rows to read
                # Pricing sheet
                pSheet = pd.read_excel(os.path.join(root, filename), sheet_name=1, header=6, names=colNames, usecols="A:H", nrows=25)
                # Intro sheet
                iSheet = pd.read_excel(os.path.join(root, filename), sheet_name=0, header=None, index_col=None, usecols="A:F", names=colNames2, skiprows=2, nrows=11)
                # Create df to fill with Intro sheet values
                current_file_intro_data = pd.DataFrame(data=iSheet)
                # Uncomment file counter below to print intro data in console to determine location of specific values
                # e.g. Rep Name is located in row 5 col F
                # print(iSheet)
                # Create df to fill with Pricing sheet values
                current_file_pricing_data = pd.DataFrame(data=pSheet)
                # Insert Intro sheet data as columns in current Pricing df based on location in Intro sheet
                # identify row, col by printing intro data df in console using commented print command above
                current_file_pricing_data.insert(7, "Rep Name", current_file_intro_data.at[5, "F"], allow_duplicates=True)
                current_file_pricing_data.insert(8, "Contact Name", current_file_intro_data.at[3, "A"], allow_duplicates=True)
                current_file_pricing_data.insert(9, "Customer Name", current_file_intro_data.at[5, "A"], allow_duplicates=True)
                current_file_pricing_data.insert(10, "Quote Number", current_file_intro_data.at[3, "F"], allow_duplicates=True)
                current_file_pricing_data.insert(11, "Promo Code", current_file_intro_data.at[4, "F"], allow_duplicates=True)
                current_file_pricing_data.insert(12, "Contact Email", current_file_intro_data.at[9, "A"], allow_duplicates=True)
                current_file_pricing_data.insert(13, "Date", current_file_intro_data.at[1, "F"], allow_duplicates=True)
                current_file_pricing_data.insert(14, "Customer Address Line 1", current_file_intro_data.at[6, "A"], allow_duplicates=True)
                current_file_pricing_data.insert(15, "Customer Address Line 2", current_file_intro_data.at[7, "A"], allow_duplicates=True)
                current_file_pricing_data.insert(16, "Customer Address Line 3", current_file_intro_data.at[8, "A"], allow_duplicates=True)
                current_file_pricing_data.insert(17, "File Name", os.path.join(root, filename), allow_duplicates=True)
                # Add combined Intro and Pricing df to new df for every file
                # so all files' data is reflected in one df
                all_data = all_data.append(current_file_pricing_data, ignore_index=True)
                # Replace empty quantity cells with "", so Pandas recognizes it is an empty string (to be later dropped)
                all_data["Quantity"].replace("", np.nan, inplace=True)
                print("replace completed")
                # Drop empty rows from df
                all_data.dropna(subset=["Quantity"], inplace=True)
                print("drop completed")
                # Uncomment file counter below to test files added to df
                # filesAdded += 1
                # print("Added: ",filesAdded,",","out of: ",totalFiles)
            # Exception handler to continue on while documenting files with errors
            except:
                print("Exception triggered: ", os.path.join(root, filename))
                continue
# Save consolidated df as csv file here: C:/Users/megan.manetas/PycharmProjects/untitled/
all_data.to_csv("Quote_Reader_2020_Test_03032020.csv")
# Test: uncomment to verify execution of code and data summary with code below
print(date_time_obj)
print("complete")
# print(all_pricing_data.head(), 11)
# print(all_pricing_data.describe())
