##### THIS FUNCTION ACCEPTS A NUMBER FROM missed_one 
##### THAT MATCHES A ROW IN THE GOOGLE SHEET.
### IT WILL REPRINT AND RESEND THE DOCX FILE IF THE 
### 24/7 SCRIPT WAS NOT RUNNING OR MISSED ONE
### TO RUN THIS SCRIPT, RUN missed_one.py

import os
import gspread
import smtplib
import dropbox
from slacker import Slacker

### IMPORT PRIVATE CONFIG FILE
import private_config as p_con

## Circular imports here, I realize ...
import create_docx as c_doc

from datetime import datetime

# Trying yagmail instead
## https://github.com/kootenpv/yagmail
import yagmail


### MY IMPORTS
#import missed_one as mo

##### SLACK TOKEN #####
#slack = Slacker(p_con.private_slack_token)

# Pulling Today's Date
now = datetime.now()
today_date = now.strftime("%m/%d/%Y %I:%M")

# Creating relative path directory
this_dir = os.path.dirname(os.path.abspath(__file__))

##### JSON Key #####
#json_key = os.path.join(this_dir, 'pcregisterpdfforwarder-1f39ce9a4ac0.json')

def email_it(row_number):

	# Initialize yagmail as yag
	yag = yagmail.SMTP(p_con.serv_email_address, p_con.serv_email_password)
	
	last_name = c_doc.worksheet.acell('B' + str(row_number)).value.strip()
	first_name = c_doc.worksheet.acell('C' + str(row_number)).value.strip()
	pet_name = c_doc.worksheet.acell('R' + str(row_number)).value.strip()
	banner_img = 'pcemailbanner.png'
	filename = os.path.join(this_dir, "submitted_customer_files/" + last_name.title().strip() + '_' + pet_name.title().strip() + ".docx")
	
	
	### SEND THE EMAIL
	subject = ("New Registration from " + pet_name.title().strip() + ' ' + last_name.title().strip() + "!")
	contents = [banner_img, first_name.title().strip() + " " + last_name.title().strip() + " has registered their dog " + pet_name.title().strip() + " with pupculture!\nThe registration form is attached to this email. \n\nIf the customer uploaded vaccination files or images, they are available in the Dropbox folder Customer Uploads. If they still need to upload these documents, they should visit pupculturenyc.com/upload.\n\n\n\nIf there is an issue with this application, please contact joshmarcus85@gmail.com\n\n", filename]
	yag.send(p_con.recipient_email, subject, contents)
	
	### MESSAGE TO ANNOUNCE EMAILING DOCUMENT
	print("\nEmailing a registration document from " + pet_name.title() + " " + 
			last_name.title() + " to " + p_con.recipient_email + "!")
		### Post Status Message to Slack Channel ###
	new_email_slack_msg = ('\nA new registration document for ' + pet_name.title() + ' ' + last_name.title() + ' has been emailed!')
	c_doc.slack.chat.post_message(p_con.slack_channel, new_email_slack_msg)
	
	################
	# Skipping the log check, as this function is for manual 
	# reprints
	'''
	print("\nEMAIL CUSTOMER REGISTRATION:")
	print("\nChecking to see if we have documents for " + 
				pet_name.title() + " " + last_name.title() +
				" already sent ...")	
	
	mail_log = open('maillog.txt', 'r')
	
	if last_name.title() and pet_name.title() in mail_log.read():
		print("\nThis customer's documents have already been emailed to pupculture! Moving on ... \n")
		pass
	else:
		print("\nEmailing a registration document from " + pet_name.title() + " " + 
			last_name.title() + " to " + p_con.recipient_email + "!")
		### Post Status Message to Slack Channel ###
		new_email_slack_msg = ('\nA new registation document for ' + pet_name.title() + ' ' + last_name.title() + ' has been emailed!')
		slack.chat.post_message(p_con.slack_channel, new_email_slack_msg)
	'''	
	#################
	
	
	#############################################
	
	'''
	##### Saving to Dropbox #####
	print("\nSaving file to Dropbox ... ")
	client = dropbox.client.DropboxClient(p_con.db_app_token)
	print 'linked account: ', client.account_info()

	os.chdir('submitted_customer_files')
				
	f = open(cust_filename, 'rb')
	response = client.put_file(('/' + last_name.title() + ', ' + pet_name.title() + ".docx"), 'f') 
	print 'uploaded ', response
		
	folder_metadata = client.metadata('/')
	print 'metadata: ', folder_metadata
		
	f, metadata = client.get_file_and_metadata('/' + last_name.title() + ', ' + pet_name.title() + ".docx")
	out = open('/' + last_name.title() + ', ' + pet_name.title() + ".docx", 'wb')
	out.write(f.read())
	out.close()
	print metadata
	'''
	return

