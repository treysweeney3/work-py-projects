#Weekly report generator for AmWaste

import pandas as pd
import PySimpleGUI as sg
import os

#Initialize GUI and take input file to pass to file_processing()
def gui_window():

    #Initialize GUI layout
    sg.theme("DefaultNoMoreNagging")
    layout = [[sg.T("")], 
              [sg.Text("Choose a file: "), sg.Input(key="-IN2-", change_submits=True), sg.FileBrowse(key="-IN-")],
              [sg.T("")], 
              [sg.Button("Submit"), sg.Button("Exit")],
              [sg.T("")],
              [sg.Text("Log: ")]]

    #Create window and handle file input
    #Assign input to df using pandas
    window = sg.Window('Report Generator for AmWaste', layout, size=(600,300))
    
    while True:
        event, values = window.read()
        #print(values["-IN2-"])
        if event == sg.WIN_CLOSED or event=="Exit":
            break
        elif event == "Submit":
            df = pd.read_excel(values["-IN-"], sheet_name=0)
            break

    return df

#Args df: takes return value of gui_window (user file input)
#Process file to generate report as specified
def file_processor(df, lookup):

    #Initialize lists for each column that will be in df2
    #Use lists to convert to DataFrame  
    date = []
    dow = []
    rt = []
    rt_type = []
    muni = []
    miles = []
    disp_tons = []
    disp_loads = []
    stops = []
    clk_hrs = []
    travel = []
    service = []
    disp = []
    pre_post = []
    target_clk_hrs = []
    variance = []

    #Iterrate through df to append to initialized lists
    #list.append: O(log^2(n)) vs df.append: O(n^2)
    #Append to lists then create df2 with completed lists
    for i, row in df.iterrows():
        date.append(str(df.at[i, 'Route Date'])[0:7])
        dow.append(str(df.at[i, 'Route Number'])[0])
        rt.append(str(df.at[i, 'Route Number'])[1:4])
        #Append route type and municipality using lookup table
        #Lookup table will be packaged with PyInstaller
        #rt_type.append() #use lookup table. figure out solution for conditionally retrieving values
        #muni.append()    #try something like: if df3.at[i, 'Route Number'] == rt: 
        for j, row in lookup.iterrows():
            if df.at[i, 'Route Number'] == lookup.at[j, 'Route']:
                rt_type.append(str(lookup.at[j, 'Route Type']))
                muni.append(str(lookup.at[j, 'Municipality']))
        miles.append(df.at[i, 'Miles'])
        disp_tons.append(df.at[i, 'Disposal Tons'])
        disp_loads.append(df.at[i, 'Disposal Loads'])
        stops.append(df.at[i, 'Stops'])
        clk_hrs.append(df.at[i, 'Clock Hours'])
        travel.append((miles[i]) / 22)
        service.append(((stops[i])^2) / 3600)
        disp.append((disp_loads[i])*(22/60))
        pre_post.append("0.78")
        target_clk_hrs.append(travel[i] + service[i] + disp[i] + 1.28)
        variance.append(target_clk_hrs[i] - clk_hrs[i])


    df2 = pd.DataFrame(list(zip(
        date, dow, rt, rt_type, muni, miles, disp_tons, disp_loads, stops, 
        clk_hrs, travel, service, disp, pre_post, target_clk_hrs, variance)), 
        columns=['Date', 'DOW', 'Route', 'Route Type', 'Municipality', 'Miles', 'Disposal Tons',
                 'Disposal Loads', 'Stops', 'Clock Hours', 'Travel', 'Service', 'Disposal', 
                 'Pre/Post', 'Target Clock Hours', 'Variance to Target'])  
    
    return df2
    
    #Add save file as to GUI
    #df2.to_excel('reporttest2.xlsx')

#Make this work lol... can't read df2
def save_file_as(df2):

    sg.theme("DefaultNoMoreNagging")
    layout = [[sg.T("")], 
              [sg.Text("Save file as: "), sg.Input(key="-IN2-", change_submits=True)],
              [sg.T("")], 
              [sg.Button("Save")]]

    #Create window and save file
    window = sg.Window('Save Report', layout, size=(600,150))
    
    while True:
        event, values = window.read()
        #print(values["-IN2-"])
        if event == sg.WIN_CLOSED or event=="Exit":
            break
        elif event == "Save":
            df2.to_excel(values["-IN2-"])
            break


#Run file_processor, passing return df from gui_window
def main():

    lookup = pd.read_excel(os.getcwd()+"/routelookup.xlsx")

    save_file_as(file_processor(gui_window(), lookup))

main()