##### This file is where you will store 
##### all private information. 
### Examples are API keys for Dropbox, Slack, 
### and anything else that uses an API key.
###
### Passwords should also be stored here and imported locally. 
###
### ADD INFORMATION ABOUT THIS TO THE README ASAP!!!

### DO NOT ADD TO GITHUB OR ANY OTHER PUBLIC SPACE!!

from slackclient import SlackClient
import os

### SLACK VARIABLES
# Helpful for running the program with no console. 
slack_channel = ['ADD SLACK CHANNEL HERE']
private_slack_token = ['ADD SLACK TOKEN HERE']

## DROPBOX TOKEN
# Not used as of yet
# dropbox_token = ['ADD DROPBOX TOKEN HERE']

### Email INformation for mail server 
# Consider adding password salting at a later date to
# better protect password information 
serv_email_address = ['ADD SERVER EMAIL ADDRESS HERE']
serv_email_password = ['ADD SERVER EMAIL PASSWORD HERE']
admin_email = ['ADD ADMIN EMAIL ADDRESS HERE']

