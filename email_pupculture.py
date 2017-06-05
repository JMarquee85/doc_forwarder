#import main_loop
## This function checks to see if the customer's registration documents
## have already been emailed to info@pupculturenyc.com by checking 
## maillog.txt. 
import os
import gspread
import smtplib
import dropbox
from oauth2client.service_account import ServiceAccountCredentials
from slacker import Slacker
import get_reg as gr

import yagmail

### IMPORT PRIVATE CONFIG FILE
import private_config as p_con
import create_docx as c_doc

from datetime import datetime

##### SLACK TOKEN #####
slack = Slacker(p_con.private_slack_token)

# Pulling Today's Date
now = datetime.now()
today_date = now.strftime("%m/%d/%Y %I:%M")

# Creating relative path directory
this_dir = os.path.dirname(os.path.abspath(__file__))

##### JSON Key #####
json_key = os.path.join(this_dir, 'pcregisterpdfforwarder-1f39ce9a4ac0.json')

def email_pupculture():
	scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
	credentials = ServiceAccountCredentials.from_json_keyfile_name(json_key, scope)
	gc = gspread.authorize(credentials)
	
	##### EMAIL PUPCULTURE FUNCTION #####	
	# See http://naelshiab.com/tutorial-send-email-python/
	reg = gc.open_by_url('https://docs.google.com/spreadsheets/d/1dYO0M9iBWVmOYcE8fO9t9rzIaTijrbLQBCNYbulAuF4/edit#gid=136975089')
	worksheet = reg.get_worksheet(0)
	pet_name = worksheet.acell('R2').value
	last_name = worksheet.acell('B2').value
	first_name = worksheet.acell('C2').value
	
	print("\nEMAIL CUSTOMER REGISTRATION:")
	print("\nChecking to see if we have documents for " + 
				pet_name.title().strip() + " " + last_name.title().strip() +
				" already sent ...")	
	
	mail_log = open('maillog.txt', 'r')
	
	if last_name.title().strip() and pet_name.title().strip() in mail_log.read():
		print("\nThis customer's documents have already been emailed to pupculture! Moving on ... \n")
		pass
	else:
		print("\nEmailing a registration document from " + pet_name.title().strip() + " " + 
			last_name.title().strip() + " to " + p_con.recipient_email + "!")
		### Post Status Message to Slack Channel ###
		new_email_slack_msg = ('\nA new registation document for ' + pet_name.title().strip() + ' ' + last_name.title().strip() + ' has been emailed!')
		slack.chat.post_message(p_con.slack_channel, new_email_slack_msg)
		
		# Initialize yagmail as yag
		yag = yagmail.SMTP(p_con.serv_email_address, p_con.serv_email_password)
	
		last_name = c_doc.worksheet.acell('B2').value.strip()
		first_name = c_doc.worksheet.acell('C2').value.strip()
		pet_name = c_doc.worksheet.acell('R2').value.strip()
		banner_img = 'pcemailbanner.png'
		filename = os.path.join(this_dir, "submitted_customer_files/" + last_name.title().strip() + '_' + pet_name.title().strip() + ".docx")
		
		### SEND THE EMAIL
		subject = ("New Registration from " + pet_name.title().strip() + ' ' + last_name.title().strip() + "!")
		contents = [banner_img, first_name.title().strip() + " " + last_name.title().strip() + " has registered their dog " + pet_name.title().strip() + " with pupculture!\nThe registration form is attached to this email. \n\nIf the customer uploaded vaccination files or images, they are available in the Dropbox folder Customer Uploads. If they still need to upload these documents, they should visit pupculturenyc.com/upload.\n\n\n\nIf there is an issue with this application, please contact joshmarcus85@gmail.com\n\n", filename]
		yag.send(p_con.recipient_email, subject, contents)
	
		
		### Post Status Message to Slack Channel ###
		new_email_slack_msg = ('\nA new registration document for ' + pet_name.title() + ' ' + last_name.title() + ' has been emailed!')
		c_doc.slack.chat.post_message(p_con.slack_channel, new_email_slack_msg)
	
	
		### Inserting file upload to Slack as a backup ###
		# This slack upload is not working for some reason
		#slack.files.upload(filename)
		
				
		# Change directory back to program root dir
		#os.chdir('..')
		
		# Log the email to a text or csv file
		# If in file, don't print message or send email	
		
		with open('maillog.txt', 'a') as mail_log:
			mail_log.write(last_name.title().strip() + ', ' + pet_name.title().strip() + ': ' + str(today_date) + "\n")
		
		'''
		##### Saving to Dropbox #####
		print("\nSaving file to Dropbox ... ")
		client = dropbox.client.DropboxClient(p_con.db_app_token)
		print 'linked account: ', client.account_info()

		os.chdir('submitted_customer_files')
				
		f = open(cust_filename, 'rb')
		response = client.put_file(('/' + last_name.title().strip() + ', ' + pet_name.title().strip() + ".docx"), 'f') 
		print 'uploaded ', response
		
		folder_metadata = client.metadata('/')
		print 'metadata: ', folder_metadata
		
		f, metadata = client.get_file_and_metadata('/' + last_name.title().strip() + ', ' + pet_name.title().strip() + ".docx")
		out = open('/' + last_name.title().strip() + ', ' + pet_name.title().strip() + ".docx", 'wb')
		out.write(f.read())
		out.close()
		print metadata
		'''
		return

