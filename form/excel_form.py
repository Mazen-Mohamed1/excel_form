import PySimpleGUI as sg
import pandas as pd

# add some color to the window
sg.theme('DarkTeal9')

excel_file = r"D:\projects\form\python_form.xlsx" #  make sure to make the excel file in the C partition to be working as administrator and to work properly
df = pd.read_excel(excel_file)

# making the layout of the form
layout = [
    [sg.Text("Please fill out the following fields: ")],
    [sg.Text("Name", size=(15, 1)), sg.InputText(key='Name')],  # this size for selecing the width and height of the cell that we are gonna write in
    [sg.Text("City", size=(15, 1)), sg.InputText(key='City')],
    [sg.Text('Favorite Color', size=(15, 1)), sg.Combo(['Green', 'Blue', 'Red'], key='Favorite Color')],  # width is 15 space and height is 1 line
    [sg.Text('I speak', size=(15, 1)),
                                        sg.Checkbox('German', key='German'),
                                        sg.Checkbox('Spanish', key='Spanish'),
                                        sg.Checkbox('English', key='English')],
    [sg.Text('No. of children',size=(15, 1)), sg.Spin([i for i in range(0, 16)],
                                                       initial_value=0, key='Children')],

    [sg.Submit(), sg.Button("Clear"),sg.Exit()]
]
def Clear_input():
    for key in values:
        window[key]('')
    return None




window = sg.Window('Simple data entry form',layout)

while True:
    event, values = window.read()
    if values and any(value == "" for value in values.values()):
        sg.popup("Please, fill out the fields")
    else:
        if event == sg.WIN_CLOSED or event == "Exit":
            break
        elif event == "Clear":
            Clear_input()
        elif event == "Submit":
            df = pd.concat([df, pd.DataFrame([values])], ignore_index=True) #  make sure to make the words in the key the same at the excel file 100%
            df.to_excel(excel_file, index=False)
            sg.popup("Data saved!")

window.close()