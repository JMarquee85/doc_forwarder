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

# EMAIL SECTION IMPORTS
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from email.MIMEBase import MIMEBase
from email.mime.application import MIMEApplication
from email.utils import formatdate
from email.utils import make_msgid
from email import encoders

### IMPORT PRIVATE CONFIG FILE
import private_config as p_con

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
				pet_name.upper() + " " + last_name.upper() +
				" already sent ...")	
	
	mail_log = open('maillog.txt', 'r')
	
	if last_name.upper() and pet_name.upper() in mail_log.read():
		print("\nThis customer's documents have already been emailed to pupculture! Moving on ... \n")
		pass
	else:
		print("\nEmailing a registration document from " + pet_name.upper() + " " + 
			last_name.upper() + " to " + p_con.recipient_email + "!")
		### Post Status Message to Slack Channel ###
		new_email_slack_msg = ('\nA new registation document for ' + pet_name.upper() + ' ' + last_name.upper() + ' has been emailed!')
		slack.chat.post_message(p_con.slack_channel, new_email_slack_msg)
		
		me = p_con.serv_email_address	# Sender
		you = p_con.recipient_email  # Recipient
		
		server = smtplib.SMTP(p_con.smtp_server, p_con.smtp_port)
		server.ehlo()
		server.starttls()
		server.ehlo()
		server.login(p_con.serv_email_address, p_con.serv_email_password)
		msg = MIMEMultipart()
		msg['Subject'] = "New Registration from " + pet_name.upper() + ' ' + last_name.upper() + "!"
		msg['From'] = me
		msg['To'] = you
 
		body = first_name.upper() + " " + last_name.upper() + " has registered their dog " + pet_name.upper() + " with pupculture!\n\n Please Note: uploaded customer vaccinations and dog images are uploaded to the Dropbox at http://pupculturenyc.com/upload"
	
		msg.attach(MIMEText(body, 'plain'))
	
		### Attach completed customer file ###
		
		filename = os.path.join(this_dir, "submitted_customer_files/" + last_name.upper() + ', ' + pet_name.upper() + ".docx")
		os.chdir('submitted_customer_files')
		cust_filename = open(last_name.upper() + ', ' + pet_name.upper() + ".docx", "rb")
		attachment = cust_filename
		#attachment = open(this_dir + "submitted_customer_files/" + last_name.upper() + ', ' + pet_name.upper() + ".docx", "rb")
	
		### Inserting file upload to Slack as a backup ###
		# This slack upload is not working for some reason
		#slack.files.upload(filename)
		
		part = MIMEBase('application', 'octet-stream')
		#part.set_payload(attachment)
		part.set_payload((attachment).read())
		encoders.encode_base64(part)
		#part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
		part.add_header('Content-Disposition', "attachment; filename= %s" % os.path.basename(filename))

		# Have noticed a bug here if there either is a 
		# docuemnt already with the filename or already listed
		# in the log files. It kills the program here if that condition
		# is happening. 
		
		msg.attach(part)
	
		text = msg.as_string()
		server.sendmail(me, you, text)
		server.quit()
		
		# Change directory back to program root dir
		os.chdir('..')
		
		# Log the email to a text or csv file
		# If in file, don't print message or send email	
		
		with open('maillog.txt', 'a') as mail_log:
			mail_log.write(last_name.upper() + ', ' + pet_name.upper() + ': ' + str(today_date) + "\n")
		
		'''
		##### Saving to Dropbox #####
		print("\nSaving file to Dropbox ... ")
		client = dropbox.client.DropboxClient(p_con.db_app_token)
		print 'linked account: ', client.account_info()

		os.chdir('submitted_customer_files')
				
		f = open(cust_filename, 'rb')
		response = client.put_file(('/' + last_name.upper() + ', ' + pet_name.upper() + ".docx"), 'f') 
		print 'uploaded ', response
		
		folder_metadata = client.metadata('/')
		print 'metadata: ', folder_metadata
		
		f, metadata = client.get_file_and_metadata('/' + last_name.upper() + ', ' + pet_name.upper() + ".docx")
		out = open('/' + last_name.upper() + ', ' + pet_name.upper() + ".docx", 'wb')
		out.write(f.read())
		out.close()
		print metadata
		'''
		return
