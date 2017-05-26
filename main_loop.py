## This file will contain functions that will receive
## the input from main_loop and or registrationtemplate,
## feed them into the pupculture registration docx template
## and then sendemail should take over from there. 
## As an optional and desired intermediary step, should also
## package document into a pdf and email. 

# http://pbpython.com/python-word-template.html

##### FUTURE GOALS FOR EXPANSION #####
# 1. Iterate through more than just first row to check for multiple 
# 		unregistered clients
# 2. Exception Handling
# 3. GUI or Command Line commands (Clear text file, etc.)
# 4. If statements to check if docx file already exists and confirm or
# deny overwrite
# 5. If no entries in Google Doc, print a message declaring this.
# 6. Show list of recent 10 entries. Allow selection of any of them 
# 		to manually recreate file. 
# 7. Mask passwords and other private info in separate private.py file
# 8. Create single variable for filepath to make changing easier
# 8b. Use tkinter dialog box to pop up and select user path selection
#		and save this information for next time
# 9. Dropbox integration
# 10. Or create separate function to call a column number or do full
#		database scan to compare logs and entire gspread doc
 

##### ISSUES LIST #####
# 1. Pulls previous registrant's information on new applications.
#		Reassignment of gspread variables at the wrong time?

from __future__ import print_function
from mailmerge import MailMerge
from datetime import datetime

import private_config as p_con
import email_pupculture as e_pc

# from slackclient import SlackClient
from slacker import Slacker

slack = Slacker(p_con.private_slack_token)

### Private configuration file


import dropbox
# See https://pypi.python.org/pypi/dropbox for more information

##### GREETING #####
print("\n\n\t-- Welcome to the pupculture Registration Forwarder --\n")
print("\t -- Written by Josh Marcus - joshmarcus85@gmail.com --\n")
print("This program creates pupculture registration forms by " +
		"pulling data from the Google Form available at register.pupculturenyc.com and emails these forms to info@pupculturenyc.com.\n\nSee the code at https://github.com/JMarquee85/pc_doc_forwarder\n")
print("\nImporting packages ... This may take a minute or so ...")


import gspread
# The below header variable not actually used here, but worth looking into later
headers = gspread.httpsession.HTTPSession(headers={'Connection':'Keep-Alive'})
import json
import os
import smtplib
import csv
from oauth2client.service_account import ServiceAccountCredentials

### POST MESSAGE TO SLACK CHANNEL

# Write this message to slack channel
# Change the channel in private_config.py
slack.chat.post_message('#pcforwarder_messages', '\n\n*** PC Forwarder Application launched! ***')

# MY IMPORTS
from get_reg import get_reg
from email_pupculture import email_pupculture
import registration_email

print("Packages imported!") 

print("\nImporting the current date ...")
print("Date imported!")

# Pulling Today's Date
now = datetime.now()
today_date = now.strftime("%m/%d/%Y %I:%M")

# Creating relative path directory
this_dir = os.path.dirname(os.path.abspath(__file__))

# JSON Credentials and Scope 

##### JSON Key #####
json_key = os.path.join(this_dir, 'pcregisterpdfforwarder-1f39ce9a4ac0.json')


##### SCOPE #####
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
# Google Credentials
credentials = ServiceAccountCredentials.from_json_keyfile_name(json_key, scope)
global gc
gc = gspread.authorize(credentials)

# Open Registration Google Spreadsheet
reg = gc.open_by_url('https://docs.google.com/spreadsheets/d/1dYO0M9iBWVmOYcE8fO9t9rzIaTijrbLQBCNYbulAuF4/edit#gid=136975089')
worksheet = reg.get_worksheet(0)
print("Google sheet successfully opened!")

##### 
# For more information on using MailMerge:
# pbpython.com/python-word-template.html
# Using templates with docx-mailmerge

# Template File
template = os.path.join(this_dir, 'pcregtemplate.docx')
#template = "/home/josh/Documents/python_work/pc_doc_forwarder/pcregtemplate.docx"

#### Do the Bit! ####

while True:
	try:
		get_reg()
		#email_pupculture()
		#registration_email()
	except KeyboardInterrupt:
		print("\nOK! Exiting program!\n") 
		### POST SLACK MESSAGE TO INFORM PROGRAM HALTED
		slack.chat.post_message('#pcforwarder_messages', '\nThe program has been exited by keyboard entry!')
		break
	except IOError:
		print("\nCreated customer not found in submitted_customer_documents directory! Trying again!")
		continue
