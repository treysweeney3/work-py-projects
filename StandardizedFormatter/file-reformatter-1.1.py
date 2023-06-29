import pandas as pd
import re
import os

#read excel file
file_path = input("Path to file to format: ")
while not os.path.isfile(file_path):
    file_path = input("File not found! Try again: ")

df = pd.read_excel(file_path)
column_header = input("Name of address column: ")

def __main__():

    user_handler()

#Defines a dictionary with street names and maps them to their corresponding abbreviations
#Uses regex 'pattern' along with str.upper() and str.replace() to update values in df[column_header]

def reformat():

    dic = {'STREET' : 'ST', 'AVENUE': 'AVE','COURT': 'CT', 'DRIVE': 'DR', 'LANE': 'LN', 
           'ROAD': 'RD', 'CIRCLE': 'CIR', 'PLACE': 'PL', 'BOULEVARD': 'BLVD', 'COVE': 'CV',
           'TERRACE': 'TER', 'PARKWAY': 'PKWY', 'RIDGE': 'RDG', 'TRAIL': 'TRL'}
    
    pattern = fr"\b({'|'.join(map(re.escape, dic.keys()))})$"
    #print(pattern)

    df[column_header] = (df[column_header].str.upper().str.replace(pattern, lambda m: dic.get(m.group()), regex=True))
    #print(df[column_header])
        
    save_name = input("Enter name to save file as **with extension**: ")
    df.to_excel(str(save_name))
    print(rf"File saved as {save_name} in C:\Users\your_name.")

    continue_formatting()

#Handles client input
#Confirms reformatting option with user

def user_handler():

    user_choice = input("Are you sure you want to reformat " + file_path + "? (Y/n): ")

    if user_choice == "y" or user_choice == "Y":
        reformat()

    elif user_choice == "n" or user_choice == "N":
        print("File will not be reformatted. Exiting program...")  
        exit()

    else:
        print("Input not recognized. Please try again.") 
        user_handler()

def continue_formatting():

    ask_continue = input("Continue? (Y/n): ")

    if ask_continue == "y" or ask_continue == "Y":
        __main__()
    elif ask_continue == "n" or ask_continue == "N":
        print("Exiting program...")
        exit()
    else:
        print("Input not recognized. Please try again.")
        continue_formatting()

__main__() 
