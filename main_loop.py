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
 
from __future__ import print_function
from requests.exceptions import ConnectionError


##### GREETING #####
print("\n\n\t-- Welcome to the pupculture Registration Forwarder --\n")
print("\t -- Written by Josh Marcus - joshmarcus85@gmail.com --\n")
print("This program creates pupculture registration forms by " +
		"pulling data from the Google Form available at register.pupculturenyc.com and emails these forms to info@pupculturenyc.com.\n\nSee the code at https://github.com/JMarquee85/doc_forwarder\n")


##### MAIN IMPORT STATEMENTS #####
from mailmerge import MailMerge
from datetime import datetime
import dropbox
# See https://pypi.python.org/pypi/dropbox for more information

##### MY IMPORTS #####
print("\nImporting packages ... This may take a minute or so ...")
### Private configuration file
import private_config as p_con
### Email PupCulture Function
#import email_pupculture as e_pc
### Get Registration Function
import get_reg as gr
### Reprint Row Function
import reprint_row as rr
###
import email_it as e_it
### 
import first_ten as ft

import gspread
import json
import sys
import os
import smtplib
import csv
import socket
import urllib, re
from oauth2client.service_account import ServiceAccountCredentials

### SLACK CLIENT ###
from slacker import Slacker
slack = Slacker(p_con.private_slack_token)

print("\nImporting the current date ...")
print("Date imported!")

# Pulling Today's Date
now = datetime.now()
today_date = now.strftime("%m/%d/%Y %I:%M")

### POST MESSAGE TO SLACK CHANNEL 

# Write this message to slack channel
# Change the channel in private_config.py
# Hostname of current computer:
this_host_hostname = socket.gethostname()
# IP Address of current computer:
this_host_ip  = re.search('"([0-9]*])"', urllib.urlopen("http://ip.jsontest.com/").read())

# Post login log to Slack channel
slack_launch_msg = ('\n** LAUNCH DETECTED ** \n' + str(now) + '\nHostname:\t' + this_host_hostname
						+ '\nIP Address: \t' + str(this_host_ip))
slack.chat.post_message(p_con.slack_channel, slack_launch_msg)


print("Packages imported!") 

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
		# Would like to start the loop automatically and 
		# interrupt it with a keypress.
		#gr.check_top_twenty()
		ft.first_ten()
		print('\n\n\nPUPCULTURE REGISTRATION MAIN MENU:\nPlease input a number ...')
		user_selection = input('\n\t1. Run the loop to check for new registrations\t' +
								'\n\t2. Change the number of registrations scanned (default 20)\t' +
								'\n\t3. Manually resend a registration (by Google Sheet row number)\t'
								'\n\t4. Clear input log and mail log\t' +
								'\n\t5. Clear all registration documents from local folder\t'+
								'\n\t6. Exit the program\t')
		
		if user_selection == 1:
			#gr.get_reg()
			gr.check_top_twenty()
			#gr.check_for_file()
			#gr.create_docx()
			#e_pc.email_pupculture()
			#registration_email()
		elif user_selection == 2:
		### Change number of registrations scanned ###
			print("\nThis feature is coming soon!")
			pass
		elif user_selection == 3:
		### Manually resend a registration form by number ###
		### missed_one.py
			rr.reprint_row()
		elif user_selection == 4:
		### Clear logs (Use carefully, as this may cause the loop to resend a bunch of emails.)
			print("\n\n*****WARNING*****\nClearing the logs may cause the program to resend a large number of emails.")
			clear_logs_query = raw_input('\nAre you sure you want to clear the log files? (y or n)\n\t')
			if clear_logs_query == 'y':
				open('inputlog.txt', 'w').close()
				open('maillog.txt', 'w').close()
				print('\n\nLogs cleared!')
			elif clear_logs_query == 'n':
				print('\n\nOk, not clearing logs. Your email inbox thanks you.')
			else:
				print('\n\nSorry, I\'m not quite sure what you meant there. Please try again.')
		elif user_selection == 5:
			### Clear all word files from folder
			print("This option is coming soon!")
		elif user_selection == 6:
			print("\nOk! Exiting program!")
			break
		else:
			print("\nSorry ... not quite sure what you meant there. Try again!")

		
	except KeyboardInterrupt:
		print("\n\nLoading main menu ...\n") 
		### POST SLACK MESSAGE TO INFORM PROGRAM HALTED
		#slack.chat.post_message(p_con.slack_channel, '\nThe program has been exited by keyboard entry!')
	#except NameError:
	#	print("\nSorry... not quite sure what you meant there. Try again, please.")
	#except SyntaxError:
	#	print("\nSorry... not quite sure what you meant there. Try again, please.")
	except IOError:
		print("\nCreated customer not found in submitted_customer_documents directory! Trying again!")
		continue
	except ConnectionError:
		print("\nUnable to connect! Please ensure you are connected to the internet! \nTrying again!")
		slack.chat.post_message(p_con.slack_channel, 'Connection lost! Attempting to reconnect ... ')
