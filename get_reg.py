#import main_loop
## pc_doc_forwarder Registration Function
# Josh Marcus
#from main_loop import gc

import gspread
import os
from oauth2client.service_account import ServiceAccountCredentials
from mailmerge import MailMerge
from datetime import datetime

## If time, or for later, it would be great to come up with a
## way to search by last name, dog name, etc. and reprint 
## rather than a row number

## Try to include that ncurses list

# Pulling Today's Date
now = datetime.now()
today_date = now.strftime("%m/%d/%Y %I:%M")
# from slackclient import SlackClient
from slacker import Slacker

import private_config as p_con
#import email_pupculture as e_pc
import import_gsheet_by_row as ig
#from email_it import *
import email_it as e_it
##### SLACK TOKEN #####
slack = Slacker(p_con.private_slack_token)

# Creating relative path directory
this_dir = os.path.dirname(os.path.abspath(__file__))

##### JSON Key #####
json_key = os.path.join(this_dir, 'pcregisterpdfforwarder-1f39ce9a4ac0.json')

from datetime import datetime

scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name(json_key, scope)
gc = gspread.authorize(credentials)
reg = gc.open_by_url('https://docs.google.com/spreadsheets/d/1dYO0M9iBWVmOYcE8fO9t9rzIaTijrbLQBCNYbulAuF4/edit#gid=136975089')
global worksheet
worksheet = reg.get_worksheet(0)

# docx Template File (pc Registration Form with MailMerge Fields)
global template
template = os.path.join(this_dir, 'pcregtemplate.docx')
# Create a mailmerge document
global document
document = MailMerge(template)


##### This for loop checks the top twenty entries
##### to see if they have already registered. 		
def check_top_twenty():
	print("\n\nREGISTRATION LOOP STARTED!")
	for x in range(p_con.search_range_x, p_con.search_range_y):
		print("\nCHECKING IF ENTRY " + str(x) + " has registered ...\n\t\t(ctrl-c to return to main menu)")
		ig.import_gsheet_by_row_and_check(x)
		


	

