# This will be a function to accept a number as an argument
# that will represent the row that will be pulled
# as a temporary stopgap to grab documents the scripy may
# have missed

# Look into option to select from a list
# with that ncurses library I ran across

from __future__ import print_function
##### MAIN IMPORT STATEMENTS #####
#from mailmerge import MailMerge
from datetime import datetime
# See https://pypi.python.org/pypi/dropbox for more information

##### MY IMPORTS #####
### Private configuration file
import private_config as p_con
import email_it as e_it
import create_docx as c_doc
import import_gsheet_by_row as ig

import sys
import os
#import smtplib
import csv
#import socket
#import urllib, re

### SLACK CLIENT ###
from slacker import Slacker
slack = Slacker(p_con.private_slack_token)

# Pulling Today's Date
now = datetime.now()
today_date = now.strftime("%m/%d/%Y %I:%M")

# Creating relative path directory
this_dir = os.path.dirname(os.path.abspath(__file__))

##### DO THE BIT! #####

while True:
	# User inputs number to select row to reprint
	try:
		# Accepts user input by row number
		print("\n\n\n\n\t\t***** PUPCULTURE REGISTRATION FORWARDER *****")
		print("\n\t\t--Manual Reprint by Google Sheet Row Number--")
		row_select = input("\nINPUT ROW NUMBER ... (2 or higher) \n\t")
		if int(row_select) <= 1:
			print("\n*** INVALID ENTRY ***\nPlease input a number of 2 or higher! \n(Note: 2 is the first row in the Google Doc.)")
			#return
		else:
			print("\nImporting customer information from Row " + 
					str(row_select) + "...")
			ig.import_gsheet_by_row(row_select)
			#d.db_upload(row_select)
			e_it.email_it(row_select)				
	except KeyboardInterrupt:
		print("\nOK! Exiting program!")
		break
	#except ConnectionError:
	#	print("\nUnable to connect! Please ensure you are connected to the internet! \nTrying again!")
	#	slack.chat.post_message(p_con.slack_channel, 'Connection lost! Attempting to reconnect ... ')
	except SyntaxError:
		print("\n\tThat didn't seem to work! Try again.\n")
	except IOError:
		print("\nFile writing error! Please make sure that you input a value of 2 or higher ...")
	except NameError:
		print("\n\tUhhhhh.... what?! Let's try that again.")
	

	
