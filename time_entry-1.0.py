#!/usr/bin/python
"""
time_entry.py - a GUI for entering and submitting time via a billing system.

Inputs:
    Billing entries.

Outputs:
    emails - a series of emails sent to the billing system properly formated for entering time.

Note:
    1. For the GUI displayed, the following applies:
        Date - The date the work was done (MM/DD/YYYY format)
        Matter No. - Self explanitory
        Rate - Your billing rate
        Timekeeper - Your timekeeper number
        Phase - The client phase code and can be blank
        Hours - Hours worked
        Description - The billing narrative
        
    2. Currently there are no checks built in to tell if data is entered incorrectly.
    	It is up to the user to make sure the date is correctly entered.

Last updated on June 11, 2021
@author: ARL

"""

## Import package/modules needed.
# Python packages/modules
import PySimpleGUI as sg
import win32com.client as client
import datetime

## Define various functions to be used.
#def submit(self):
def submit(DATE, MATTER, HOURS, PHASE, TK_NUM, RATE, DESCRIPTION):
    # Initialize variables needed.
    to_email = 'user@email.com'       # development email
                
    # Concatenate the various strings into a single string for each item of an entry.
    dateDATE = 'INCUR DATE=' + DATE
    hoursHOURS = 'HOURS=' + HOURS
    discDISC = 'DESCRIPTION=' + DESCRIPTION
    rateRATE = 'Rate=' + RATE
    time_kTIME_K = 'Timekeeper=' + TK_NUM
    phasePHASE = 'Phase=' + PHASE
    msg = f"{dateDATE}\n{hoursHOURS}\n{rateRATE}\n{time_kTIME_K}\n{phasePHASE}\n{discDISC}\n"
    
    # Acually send the email.
    outlook = client.Dispatch("Outlook.Application")
    message = outlook.CreateItem(0)
    message.To = to_email
    message.Subject = MATTER
    message.Body = msg
    message.Send()
# end submit function

def clear():
    #window['-DATE-'].update('')
    window['-MATTER-'].update('')
    window['-HOURS-'].update('')
    window['-PHASE-'].update('')
    #window['-TK_NUM-'].update('')
    #window['-RATE-'].update('')
    window['-DESCRIPTION-'].update('')
# end clear function

#DATE = datetime.date.today()
#DATE = DATE.strftime("%m/%d/%Y")

sg.theme('Dark Blue 3')   # Add a touch of color

layout = [[sg.Text('Time Entry System', justification='')],
          [sg.Text('Date:', size=(13, 1), justification='right'), sg.InputText(key='-DATE-')],
          [sg.Text('Matter No.:', size=(13, 1), justification='right'), sg.InputText(key='-MATTER-')],
          [sg.Text('Hours:', size=(13, 1), justification='right'), sg.InputText(key='-HOURS-')],
          [sg.Text('Phase:', size=(13, 1), justification='right'), sg.InputText(key='-PHASE-')],
          [sg.Text('Time Keeper No.:', size=(13, 1), justification='right'), sg.InputText(key='-TK_NUM-')],
          [sg.Text('Rate:', size=(13, 1), justification='right'), sg.InputText(key='-RATE-')],
          [sg.Text('Description:', size=(13, 1), justification='right'), sg.Multiline(key='-DESCRIPTION-', size=(43,5))],
          [sg.Button('Submit'), sg.Button('Clear'), sg.Button('Exit')]]


window = sg.Window('Time Entry System', layout)
# Event Loop to process "events" and get the "values" of the inputs
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit': # if user closes window or clicks cancel
        break
    elif event == 'Clear':
        clear()
    else:
        submit(values['-DATE-'], values['-MATTER-'], values['-HOURS-'], values['-PHASE-'], values['-TK_NUM-'], values['-RATE-'], values['-DESCRIPTION-'])        
        clear()
    # end, if elif, else.

window.close()

# to build executable use "auto-py-to-exe"
