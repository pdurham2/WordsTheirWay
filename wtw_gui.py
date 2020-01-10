import PySimpleGUI as sg      


sg.theme('DarkAmber')    # Remove line if you want plain gray windows

layout = [ 
            [sg.Text('Select Feature Guide')],
            # [sg.Listbox(values=('primary', 'elementary', 'upper-level'), size=(20, 3))],
            [sg.Radio('Primary', "feature_guide_radio", default=True, size=(10,1)), sg.Radio('Elementary', "feature_guide_radio"), sg.Radio('Upper Level', "feature_guide_radio")],
            [sg.Text('Document to score')],
            [sg.In(), sg.FileBrowse()],
#            [sg.Open(), sg.Cancel()],      
            [sg.Button('Run', bind_return_key=True), sg.Cancel()]      
            ]
window = sg.Window('Words Their Way', layout)      

while True:      
    event, values = window.read()      
    if event is not None and event not in ('Cancel'):      
        try:  
            if values[0]:
                selection = "primary"
            elif values[1]:
                selection = "elementary"
            else:
                selection = "upper-level"


            print("Feature Guide: ", selection)
            print("File path: ", values[3], values["Browse"])  
        except:      
            print("Error")
        # window['output'].update("Scored workbook saved here:")      
    else:      
        break   
