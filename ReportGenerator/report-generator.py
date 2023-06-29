#Weekly report generator for AmWaste

import pandas as pd
import PySimpleGUI as sg

#Initialize GUI and take input file to pass to file_processing()
def gui_window():

    #Initialize GUI layout
    sg.theme("DefaultNoMoreNagging")
    layout = [[sg.T("")], [sg.Text("Choose a file: "), sg.Input(key="-IN2-", change_submits=True),
                    sg.FileBrowse(key="-IN-")], [sg.T("")], [sg.Button("Submit"), sg.Button("Exit")]]

    #Create window and handle file input
    #Assigns input to df using pandas
    window = sg.Window('Report Generator for AmWaste', layout, size=(600,150))
    
    while True:
        event, values = window.read()
        print(values["-IN2-"])
        if event == sg.WIN_CLOSED or event=="Exit":
            break
        elif event == "Submit":
            df = pd.read_excel(values["-IN-"])
            break

    return df

#Args df: takes return value of gui_window (user file input)
#Processes file to generate report as specified
def file_processor(df):

    print(df.head())

def main():

    file_processor(gui_window())

main()