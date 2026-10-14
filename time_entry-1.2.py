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

Last updated on June 17, 2021
@author: ARL

"""

## Import package/modules needed.
# Python packages/modules
import PySimpleGUI as sg
import win32com.client as client
import datetime
from configparser import ConfigParser
import json

## Define various functions to be used.
# Function to generate collect data and format billing entry
def entry(DATE, HOURS, PHASE, TK_NUM, RATE, DESCRIPTION):
    # Concatenate the various strings into a single string for each item of an entry.
    dateDATE = 'INCUR DATE=' + DATE
    hoursHOURS = 'HOURS=' + HOURS
    discDISC = 'DESCRIPTION=' + DESCRIPTION
    rateRATE = 'Rate=' + RATE
    time_kTIME_K = 'Timekeeper=' + TK_NUM
    phasePHASE = 'Phase=' + PHASE
    msg = f"{dateDATE}\n{hoursHOURS}\n{rateRATE}\n{time_kTIME_K}\n{phasePHASE}\n{discDISC}\n"
    return (msg)
# end entry function

# Function to send emails
def email_func(to_email, subject, msg):
    try:
        # Acually send the email
        outlook = client.Dispatch("Outlook.Application")
        message = outlook.CreateItem(0)
        message.To = to_email
        message.Subject = subject
        message.Body = msg
        message.Send()
        return True
    except:
        return False
# end submit function

# Function to clear data fields after sending entry
def clear():
    #window['-DATE-'].update('')
    window['-MATTER-'].update('')
    window['-HOURS-'].update('')
    window['-PHASE-'].update('')
    #window['-TK_NUM-'].update('')
    #window['-RATE-'].update('')
    window['-DESCRIPTION-'].update('')
# end clear function

# Function to load config data
def load_config():
    try:
        config = ConfigParser()
        config.read('config.ini')
        user_email = config['user']['email']
        billing = config['user']['billing_rate']
        TK_NUM = config['user']['tk_number']
        multi_bill = config['user']['multi_bill_rate']
        user_data = (user_email, billing, TK_NUM, multi_bill)
        return user_data
    except:
        sg.popup("Configuration file not found.  Please make sure 'config.int' file is located in same directory as the executable file and try again.", title='Config File Error')
        return False
# end load_config function

# Function to extract standard billing from dictionary.
def stand_bill(disc, bill_strings):   
    disc = disc[:-1] 
    disc = bill_strings.get(disc)
    if disc == None:   
        return False
    else:
        return disc
# end stand_bill function


## Actual Program ##
# Initialize variables needed
to_email = 'user@email.com'
user_data = load_config()           # user data from config file
tk_number = user_data[2]            # set user's timekeeper number
DATE = datetime.date.today()        # set default date to current day
DATE = DATE.strftime("%m/%d/%Y")    # format date
with open('bill_strings.json') as f: # load dictionary of standard billing entries for user.
    bill_strings = json.loads(f.read())
    f.close()

# Set layout for window
sg.theme('Dark Blue 3')   # Add a touch of color

if user_data[3] == 'yes': # Set layout for user with multiple billing rates
    layout = [[sg.Text('Time Entry System', justification='')],
              [sg.Text('Date:', size=(13, 1), justification='right'), sg.InputText(DATE, key='-DATE-')],
              [sg.Text('Matter No.:', size=(13, 1), justification='right'), sg.InputText(key='-MATTER-')],
              [sg.Text('Hours:', size=(13, 1), justification='right'), sg.InputText(key='-HOURS-')],
              [sg.Text('Phase:', size=(13, 1), justification='right'), sg.InputText(key='-PHASE-')],              
              [sg.Text('Rate:', size=(13, 1), justification='right'), sg.InputText(user_data[1], key='-RATE-')],
              [sg.Text('Description:', size=(13, 1), justification='right'), sg.Multiline(key='-DESCRIPTION-', size=(43,5))],
              [sg.Button('Submit'), sg.Button('Clear'), sg.Button('Exit')]]
else:
    layout = [[sg.Text('Time Entry System', justification='')],
              [sg.Text('Date:', size=(13, 1), justification='right'), sg.InputText(DATE, key='-DATE-')],
              [sg.Text('Matter No.:', size=(13, 1), justification='right'), sg.InputText(key='-MATTER-')],
              [sg.Text('Hours:', size=(13, 1), justification='right'), sg.InputText(key='-HOURS-')],
              [sg.Text('Phase:', size=(13, 1), justification='right'), sg.InputText(key='-PHASE-')],
              [sg.Text('Description:', size=(13, 1), justification='right'), sg.Multiline(key='-DESCRIPTION-', size=(43,5))],
              [sg.Button('Submit'), sg.Button('Clear'), sg.Button('Exit')]]
    rate = user_data[1]

## Define window variable and open window.
window = sg.Window('Time Entry System', layout)

while True: # Event Loop to process "events" and get the "values" of the inputs
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit': # if user closes window or clicks cancel
        break
    elif event == 'Clear':
        clear()  #clear data fields
        if user_data[3] == 'yes':
            window['-RATE-'].update(user_data[1]) # rest billing rate to default rate        
    else:        
        if bill_strings != False:
            temp = stand_bill(values['-DESCRIPTION-'], bill_strings)            
            if temp == False:
                disc = values['-DESCRIPTION-']                
            else:
                disc = temp
            # end if, else
        else:
            disc = values['-DESCRIPTION-']
        # end if, else
        
        # Prepare data for the email
        if user_data[3] == 'yes':   # creat message for multi billing rate user         
            msg = entry(values['-DATE-'], values['-HOURS-'], values['-PHASE-'], tk_number, values['-RATE-'], disc)
        else:  # create message for single billing rate user
            msg = entry(values['-DATE-'], values['-HOURS-'], values['-PHASE-'], tk_number, rate, disc)

        subject = values['-MATTER-']
        email_func(to_email, subject, msg) # actually send the email
        clear() # clear data fields
        if user_data[3] == 'yes':
            window['-RATE-'].update(user_data[1]) # reset billing rate to default rate
        
    # end, if elif, else.

window.close()

# to build executable use "auto-py-to-exe"
