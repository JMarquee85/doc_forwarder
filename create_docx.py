##### THIS FILE CREATES A DOCX FILE FROM THE 
### GOOGLE SHEET AND SAVES IT TO BE EMAILED OR ACCESSED LATER

import gspread
import os
from oauth2client.service_account import ServiceAccountCredentials
from slacker import Slacker
from mailmerge import MailMerge
import private_config as p_con
from datetime import datetime

### SLACK CLIENT ###
from slacker import Slacker
slack = Slacker(p_con.private_slack_token)

# Pulling Today's Date
now = datetime.now()
today_date = now.strftime("%m/%d/%Y %I:%M")

# Creating relative path directory
this_dir = os.path.dirname(os.path.abspath(__file__))

# JSON Credentials and Scope 
##### JSON Key #####
### Move this to private_config soon. 
json_key = os.path.join(this_dir, 'pcregisterpdfforwarder-1f39ce9a4ac0.json')

##### SCOPE #####
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
# Google Credentials
credentials = ServiceAccountCredentials.from_json_keyfile_name(json_key, scope)
global gc
gc = gspread.authorize(credentials)

# Open Registration Google Spreadsheet
reg = gc.open_by_url('https://docs.google.com/spreadsheets/d/1dYO0M9iBWVmOYcE8fO9t9rzIaTijrbLQBCNYbulAuF4/edit#gid=136975089')
global worksheet
worksheet = reg.get_worksheet(0)

##### 
# For more information on using MailMerge:
# pbpython.com/python-word-template.html
# Using templates with docx-mailmerge

# Template File
template = os.path.join(this_dir, 'pcregtemplate.docx')

##### CREATE DOCX FILE #####


##### CREATE DOC #####
def create_docx(row_number):

		
	#### CREATE THE DOCUMENT #####

	print("\nCreating customer document for " + pet_name.title().strip() + 
			" " + last_name.title().strip() + " ...")
	document.merge(
		last_name = last_name,
		first_name = first_name,
		address = street,
		apartment = apt,
		cross_streets = cross_streets,
		city = city,
		state = state, 
		zip = zip_code,
		home_phone = home_phone,
		cell_phone = cell_phone,
		business_phone = business_phone,
		email_address = email_address,
		reference = reference,
		emergency_contacts = emergency_contacts,
		membership = membership,
		keys = keys,
		pet_name = pet_name,
		pet_nick = pet_nick,
		breed = breed,
		weight = weight,
		sex = sex,
		dob = dob,
		color = color,
		spayed = fixed,
		how_long_owned = length_owned,
		brand_food = brand_food,
		times_fed = times_per_day_feed,
		size_portion = serving_size,
		allergies_digestive = allergies_digestive,
		treats_ok = treats_ok,
		fed_during_daycare = fed_during_daycare,
		in_home_restrictions = home_restriction,
		water_out = water_out,
		dry_food = dry_food, 
		leash_location = leash_location,
		where_is_the_stuff = where_is_the_stuff,
		where_is_the_other_stuff = where_is_the_other_stuff,
		lights_locks = lights_locks,
		quirks = helpful_info,
		dog_allowed = dog_allowed,
		behavior = behavior,
		behavior_new_people = behavior_new_people,
		behavior_leash = behavior_leash,
		housebroken = housebroken,
		last_vet = last_vet,
		flea_prevention = flea_prevention,
		medical_conditions = medical_conditions,
		medication = medications,
		ever_bitten = ever_attacked,
		vet_name = vet_name,
		vet_phone = vet_phone,
		vet_address = vet_address,
		vet_city = vet_city,
		vet_state = vet_state,
		date = today_date)
		# Write completed document
		# First, change directory to submitted_customer_files
	print("Writing new customer file ...")
	os.chdir('submitted_customer_files')
	document.write(last_name.title().strip() + '_' + pet_name.title().strip() + '.docx')
	os.chdir('..')
			# Terminal Message to Confirm Success
	new_reg_msg = "\nA registration docx form for " + pet_name.title().strip() + " " + last_name.title().strip() + " has been created!\t"
	print(new_reg_msg)
	
	### Post Slack Message to Update Channel
	slack.chat.post_message(p_con.slack_channel, new_reg_msg)
	
	##### LOGGING TO TEXT FILE #####
	with open('inputlog.txt', 'a') as doglog:
		doglog.write('\n' + last_name.title().strip() + ', ' + first_name.title().strip() + ' (Owner), ' + pet_name.title().strip() + ':  ' + str(today_date) + '  --MANUAL REPRINT') 
	### Confirm Logging ###
	print("Registration information for " + pet_name.title().strip() + " " + last_name.title().strip() + 
				" has been logged to inputlog.txt!")
		
	return
