import pandas as pd
import re

#read excel file
file_path = input("Path to file to format: ")
df = pd.read_excel(file_path)
column_header = input("Name of address column: ")

def __main__():

    search = input("Enter part of a string that you would like to be reformatted (ex. aven, pla, st, etc.): ")
    format_to = input("Enter what this should be reformatted to: ")

    user_handler(search, format_to)

#Takes args search(str) and format_to(str)
#Iterates through rows in user specified column to search for search var
#If found, replaces value in current cell with format_to

def reformat(search, format_to):

    count = 0

    for i, row in df.iterrows():
        cell_value = df.at[i, column_header]
        #update search with regex expression
        #try an exp that will partition string and find substring at end of string
        if cell_value == str.match(search, case=False):
            cell_value = str(format_to)
            count += 1
        else:
            print("No matches found...\nReturning to search...")
            __main__()
        df.at[i, column_header] = cell_value
        
    print(str(count) + " address(es) have been updated.")
    df.to_excel("test2.xlsx")

    continue_formatting()

#Takes args search(str) and format_to(str) in order to reuse reformat() call
#Handles client input
#Confirms reformatting option with user

def user_handler(search, format_to):

    user_choice = input("Are you sure you want to reformat " + file_path + "? (Y/n) ")

    if user_choice == "y" or user_choice == "Y":
        reformat(search, format_to)

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
