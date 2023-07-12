#Weekly report generator for AmWaste

import pandas as pd
import PySimpleGUI as sg
import os
import sys

#Initialize GUI and take input file to pass to file_processing()
def gui_window():

    #Initialize GUI layout
    sg.theme("DefaultNoMoreNagging")
    layout = [[sg.T("")], 
              [sg.Text("Choose a file: "), sg.Input(key="-IN2-", change_submits=True), sg.FileBrowse(key="-IN-")],
              [sg.T("")], 
              [sg.Button("Submit"), sg.Button("Exit")]]

    #Create window and handle file input
    #Assign input to df using pandas
    window = sg.Window('Report Generator for AmWaste', layout, size=(550,150))
    
    while True:
        event, values = window.read()
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

    #Iterate through df to append to initialized lists
    #list.append: O(log^2(n)) vs df.append: O(n^2)
    #Append to lists then create df2 with completed lists
    for i, row in df.iterrows():
        date.append(str(df.at[i, 'Route Date'])[0:10])
        dow.append(str(df.at[i, 'Route Number'])[0])
        rt.append(str(df.at[i, 'Route Number'])[1:4])
        #Append route type and municipality using lookup table
        #Lookup table will be packaged with PyInstaller
        for j, row in lookup.iterrows():
            if df.at[i, 'Route Number'] == lookup.at[j, 'Route']:
                rt_type.append(str(lookup.at[j, 'Route Type']))
                muni.append(str(lookup.at[j, 'Municipality']))
            else:
                continue
        miles.append(df.at[i, 'Miles'])
        disp_tons.append(df.at[i, 'Disposal Tons'])
        disp_loads.append(df.at[i, 'Disposal Loads'])
        stops.append(df.at[i, 'Stops'])
        clk_hrs.append("{:.2f}".format(df.at[i, 'Clock Hours']))
        travel.append("{:.2f}".format((miles[i]) / 22))
        #Service time varies by truck type and municipality
        #Could've added another lookup table for this, but hard coded implementation was faster here
        if rt_type[i] == 'AFEL' and muni[i] == 'Vestavia Hills':
            service.append("{:.2f}".format((stops[i]*17)/3600))
        elif rt_type[i] == 'AFEL' and muni[i] == 'Jefferson County':
            service.append("{:.2f}".format((stops[i]*14)/3600))
        elif rt_type[i] == 'PacMac' and muni[i] == 'Vestavia Hills':
            service.append("{:.2f}".format((stops[i]*27)/3600))
        elif rt_type[i] == 'PacMac' and muni[i] == 'Jefferson County':
            service.append("{:.2f}".format((stops[i]*28)/3600))
        elif rt_type[i] == 'ParKan' and muni[i] == 'Vestavia Hills':
            service.append("{:.2f}".format((stops[i]*17)/3600))
        elif rt_type[i] == 'ParKan' and muni[i] == 'Jefferson County':
            service.append("{:.2f}".format((stops[i]*28)/3600))
        elif rt_type[i] == 'AFEL' and muni[i] == 'Hoover':
            service.append("{:.2f}".format((stops[i]*17)/3600))
        elif rt_type[i] == 'PacMac' and muni[i] == 'Hoover':
            service.append("{:.2f}".format((stops[i]*17)/3600))
        elif rt_type[i] == 'REL' and muni[i] == 'Fultondale':
            service.append("{:.2f}".format((stops[i]*18)/3600))
        elif rt_type[i] == 'REL' and muni[i] == 'Jefferson County':
            service.append("{:.2f}".format((stops[i]*18)/3600))
        elif rt_type[i] == 'AFEL' and muni[i] == 'Pelham':
            service.append("{:.2f}".format((stops[i]*17)/3600))
        elif rt_type[i] == 'AFEL' and muni[i] == 'Pelham':
            service.append("{:.2f}".format((stops[i]*17)/3600))
        elif rt_type[i] == 'PacMac' and muni[i] == 'Pelham':
            service.append("{:.2f}".format((stops[i]*17)/3600))
        elif rt_type[i] == 'REL' and muni[i] == 'Mountain Brook':
            service.append("{:.2f}".format((stops[i]*18)/3600))
        elif rt_type[i] == 'AFEL' and muni[i] == 'Mountain Brook':
            service.append("{:.2f}".format((stops[i]*17)/3600))
        elif rt_type[i] == 'PacMac' and muni[i] == 'Mountain Brook':
            service.append("{:.2f}".format((stops[i]*39)/3600))
        elif rt_type[i] == 'ASL' and muni[i] == 'Jefferson County':
            service.append("{:.2f}".format((stops[i]*14)/3600))
        elif rt_type[i] == 'AFEL' and muni[i] == 'Trussville':
            service.append("{:.2f}".format((stops[i]*14)/3600))
        elif rt_type[i] == 'ASL' and muni[i] == 'Trussville':
            service.append("{:.2f}".format((stops[i]*14)/3600))
        elif rt_type[i] == 'PacMac' and muni[i] == 'Trussville':
            service.append("{:.2f}".format((stops[i]*17)/3600))
        elif rt_type[i] == 'ParKan' and muni[i] == 'Mountain Brook':
            service.append("{:.2f}".format((stops[i]*39)/3600))
        else:
            service.append('0')
        if disp_loads[i] != None:
            disp.append("{:.2f}".format((disp_loads[i])*(22/60)))
        else:
            disp.append(int(''))
        pre_post.append("0.78")
        target_clk_hrs.append("{:.2f}".format(float(travel[i]) + float(service[i]) + float(disp[i]) + 1.28))
        variance.append("{:.2f}".format(float(target_clk_hrs[i]) - float(clk_hrs[i])))

    df2 = pd.DataFrame(list(zip(
        date, dow, rt, rt_type, muni, miles, disp_tons, disp_loads, stops, 
        clk_hrs, travel, service, disp, pre_post, target_clk_hrs, variance)), 
        columns=['Date', 'DOW', 'Route', 'Route Type', 'Municipality', 'Miles', 'Disposal Tons',
                 'Disposal Loads', 'Stops', 'Clock Hours', 'Travel', 'Service', 'Disposal', 
                 'Pre/Post', 'Target Clock Hours', 'Variance to Target']) 
    
    def stylevariance(val):
        color = 'black' if str(val) == 'nan' else 'red' if str(val) < str(0) else 'green' if str(val) > str(0) else ''
        return 'background-color: %s' % color
    
    styled_df2 = df2.style
    styled_df2 = styled_df2.applymap(lambda x: 'background-color: yellow', subset=['Travel', 'Service', 'Disposal', 'Pre/Post', 'Target Clock Hours'])
    styled_df2 = styled_df2.applymap(stylevariance, subset=['Variance to Target'])

    return styled_df2

def save_file_as(df2):

    sg.theme("DefaultNoMoreNagging")
    layout = [[sg.T("")], 
              [sg.Text("Save file as: "), 
              sg.Input(key="-IN2-", change_submits=True), sg.FileSaveAs(key="-IN-", 
                file_types=(('Excel Workbook', '.xlsx'), ('Excel 97-2003 Workbook', '.xls')))],
              [sg.T("")], 
              [sg.Button("Save")]]

    #Create window and save file
    window = sg.Window('Save Report', layout, size=(600,150))
    
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event=="Exit":
            break
        elif event == "Save":
            df2.to_excel(values["-IN-"])
            break

def main():

    #For executable use
    lookup = pd.read_excel(os.path.abspath(sys._MEIPASS + "\\resources\\routelookup.xlsx"))
    #For nonexecutable use only
    #lookup = pd.read_excel("C:\\Users\\tsweeney\\Documents\\work-py-projects\\ReportGenerator\\resources\\routelookup.xlsx")

    save_file_as(file_processor(gui_window(), lookup))

main()

