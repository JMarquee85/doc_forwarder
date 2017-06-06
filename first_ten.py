import gspread
import os
from oauth2client.service_account import ServiceAccountCredentials
from mailmerge import MailMerge


# from slackclient import SlackClient
from slacker import Slacker

import private_config as p_con
import email_pupculture as e_pc
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
worksheet = reg.get_worksheet(0)
# Pulling Today's Date
now = datetime.now()
today_date = now.strftime("%m/%d/%Y %I:%M")
	
# docx Template File (pc Registration Form with MailMerge Fields)
template = os.path.join(this_dir, 'pcregtemplate.docx')
# Create a mailmerge document
document = MailMerge(template)	

# Last, First, Pet (B, C, R)
cust_1 = [worksheet.acell('B2').value.strip(), worksheet.acell('C2').value.strip(), worksheet.acell('R2').value.strip()]
cust_2 = [worksheet.acell('B3').value.strip(), worksheet.acell('C3').value.strip(), worksheet.acell('R3').value.strip()]
cust_3 = [worksheet.acell('B4').value.strip(), worksheet.acell('C4').value.strip(), worksheet.acell('R4').value.strip()]
cust_4 = [worksheet.acell('B5').value.strip(), worksheet.acell('C5').value.strip(), worksheet.acell('R5').value.strip()]
cust_5 = [worksheet.acell('B6').value.strip(), worksheet.acell('C6').value.strip(), worksheet.acell('R6').value.strip()]
cust_6 = [worksheet.acell('B7').value.strip(), worksheet.acell('C7').value.strip(), worksheet.acell('R7').value.strip()]
cust_7 = [worksheet.acell('B8').value.strip(), worksheet.acell('C8').value.strip(), worksheet.acell('R8').value.strip()]
cust_8 = [worksheet.acell('B9').value.strip(), worksheet.acell('C9').value.strip(), worksheet.acell('R9').value.strip()]
cust_9 = [worksheet.acell('B10').value.strip(), worksheet.acell('C10').value.strip(), worksheet.acell('R10').value.strip()]
cust_10 = [worksheet.acell('B11').value.strip(), worksheet.acell('C11').value.strip(), worksheet.acell('R11').value.strip()]
custs = [
	cust_1,
	cust_2,
	cust_3,
	cust_4,
	cust_5,
	cust_6,
	cust_7,
	cust_8,
	cust_9,
	cust_10
]

def first_ten():
	#for customers in custs:
	#	print customers
	print("\nLATEST 10 REGISTRATIONS:\n")	
	print("\t1. " + custs[0][2].title().strip() + " " + custs[0][0].title().strip())
	print("\t2. " + custs[1][2].title().strip() + " " + custs[1][0].title().strip())
	print("\t3. " + custs[2][2].title().strip() + " " + custs[2][0].title().strip())
	print("\t4. " + custs[3][2].title().strip() + " " + custs[3][0].title().strip())
	print("\t5. " + custs[4][2].title().strip() + " " + custs[4][0].title().strip())
	print("\t6. " + custs[5][2].title().strip() + " " + custs[5][0].title().strip())
	print("\t7. " + custs[6][2].title().strip() + " " + custs[6][0].title().strip())
	print("\t7. " + custs[7][2].title().strip() + " " + custs[7][0].title().strip())
	print("\t9. " + custs[8][2].title().strip() + " " + custs[8][0].title().strip())
	print("\t10. " + custs[9][2].title().strip() + " " + custs[9][0].title().strip())
	
#first_ten()
		
		
	


