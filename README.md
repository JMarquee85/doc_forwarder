# Doc Forwarder
A very simple backend application to turn Google Form spreadsheet fields
into docx files and email them. <br><br>

Project written for http://register.pupculturenyc.com<br>

 Written by Josh Marcus <br>
 joshmarcus85@gmail.com <br>
 http://jmarquee.xyz <br>

# WHAT IT DOES: <br>
 - Watches a Google Spreadsheet for changes 
 - Takes user input of application form received from Google Form and Google Sheet and creates
                                 a mailmerge docx based on an existing template (pupculture registration form). 
 - Emails said document(s) to company site
 - (Coming soon!!) Stores documents in Dropbox and/ or system folder
 - Emails customer based on conditional response (if they would like to receive more information about a topic)

# RUNNING IT: <br>

 Make sure you have all the required packages installed by running:
	```pip install -r requirements.txt```
 Change directory to the project:
    
    ```cd /path/to/directory```
 Then start the program with the following command:
    ```python main_loop.py```

# CONFIGURATION: <br>
 WARNING: DO NOT store the variables in this file publicly!!
 
 - Rename private_config_example.py to private_config.py
 - Get <a href="https://api.slack.com/custom-integrations/legacy-tokens"> LEGACY TOKEN </a>
 from Slack to post status messages to a slack channel. Store your 
 Slack token in your private_config.py file as private_slack_token. 
 - Store your Slack Channel name in your private_config.py 
 in slack_channel. 
 - (DROPBOX non-functional as of this writing!) Set up an app at 
 <a href="https://api.slack.com/custom-integrations/legacy-tokens">Dropbox</a> and store the 
 app token, app key, and secret key in your private_config.py file. 
 - Add the sender email address under serv_email_address in private_config.py. 
 Store the email password in serv_email password. (admin_email not currently used.)

# STOPPING IT: <br> 
Simply type ```ctrl-c``` to quit the program.



