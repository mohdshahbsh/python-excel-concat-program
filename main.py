"""
Excel Concatonation Program

Dependencies:
1. pip install openpyxl
2. pip install pandas
"""
import os
import pandas as pd
import numpy as np
import glob
import re


def read_excel(file_name):
    """This function is used to read the excel file"""
    return pd.read_excel(file_name)


def read_index(index_str):
    """This function is used to apply the index logic to the input"""
    # remove whitespaces in the string
    clean_index_str = [index.strip() for index in re.split(',', index_str)]
    
    # apply range logic
    for i, index in enumerate(clean_index_str):
        if index.isdigit():
            clean_index_str[i] = int(index)

        elif '-' in index:
            index_range = [int(n) for n in index.split('-')]
            index_list = [x for x in range(index_range[0], index_range[1]+1)]

            clean_index_str.pop(i)
            clean_index_str = clean_index_str + index_list

    return clean_index_str


def main():
    # 1. Read all the Excel files in the directory
    file_list = [file for file in glob.glob(".\*.xlsx")]

    print(f"\nExcel files that exist in the directory {os.getcwd()}\n")
    for i, file_name in enumerate(file_list):
        print(f"    {i:2}. {file_name}")
    
    # 2. Request file index to concat and file name to save as
    concat_index = input("\nSelect files to union (i.e. 1,3-5): ")
    concat_file_name = input("Output file name: ")

    # Error handling for wrong input parameter
    for input_index in re.split(',|-', concat_index):
        if not input_index.strip().isdigit():
            raise Exception("Error: You did not insert a proper index number, please insert correct index")

        elif not int(input_index) < len(file_list):
            raise Exception(f"Error: The index {input_index} is out of range")

    index_list = read_index(concat_index)

    # 3. Concat the tables and save
    frames_list = [read_excel(file_list[i]) for i in index_list]
    concat_frame = pd.concat(frames_list, ignore_index=True)

    concat_frame.to_excel(f"{concat_file_name}.xlsx", index=False)

    print("\nThe Excel files has been successfully combined")



if __name__ == "__main__":
    main()
