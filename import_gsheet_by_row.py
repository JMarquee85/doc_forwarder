
import gspread
import os
from oauth2client.service_account import ServiceAccountCredentials
from slacker import Slacker
from mailmerge import MailMerge
import private_config as p_con
from datetime import datetime
import create_docx as c_doc

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

def import_gsheet_by_row(row_number):
	
	# docx Template File (pc Registration Form with MailMerge Fields)
	template = os.path.join(this_dir, 'pcregtemplate.docx')
	# Create a mailmerge document and accept the template file
	document = MailMerge(template)	
	
	##### VARIABLES SECTION #####
	### Import variables from Google Sheet
	time_registered = worksheet.acell('A' + str(row_number)).value.strip()
	last_name = worksheet.acell('B' + str(row_number)).value.strip()
	first_name = worksheet.acell('C' + str(row_number)).value.strip()
	street = worksheet.acell('D' + str(row_number)).value.strip()
	apt = worksheet.acell('E' + str(row_number)).value.strip()
	cross_streets = worksheet.acell('F' + str(row_number)).value.strip()
	city = worksheet.acell('G' + str(row_number)).value.strip()
	state = worksheet.acell('H' + str(row_number)).value.strip()
	zip_code = worksheet.acell('I' + str(row_number)).value.strip()
	home_phone = worksheet.acell('J' + str(row_number)).value.strip()
	cell_phone = worksheet.acell('K' + str(row_number)).value.strip()
	business_phone = worksheet.acell('L' + str(row_number)).value.strip()
	email_address = worksheet.acell('M' + str(row_number)).value.strip()
	# Where did you hear about us?
	reference = worksheet.acell('N' + str(row_number)).value.strip()
	emergency_contacts = worksheet.acell('O' + str(row_number)).value.strip()
	membership = worksheet.acell('P' + str(row_number)).value.strip()
	keys = worksheet.acell('Q' + str(row_number)).value.strip()
	# Pet Information Section
	pet_name = worksheet.acell('R' + str(row_number)).value.strip()
	pet_nick = worksheet.acell('S' + str(row_number)).value.strip()
	breed = worksheet.acell('T' + str(row_number)).value.strip()
	weight = worksheet.acell('U' + str(row_number)).value.strip()
	sex = worksheet.acell('V' + str(row_number)).value.strip()
	dob = worksheet.acell('W' + str(row_number)).value.strip()
	color = worksheet.acell('X' + str(row_number)).value.strip()
	fixed = worksheet.acell('Y' + str(row_number)).value.strip()
	# How long have you owned your dog?
	length_owned = worksheet.acell('Z' + str(row_number)).value.strip()
	# Please specify what brand and type (wet/ dry) food you use:
	brand_food = worksheet.acell('AA' + str(row_number)).value.strip()
	# How many times per day do you feed your dog?
	times_per_day_feed = worksheet.acell('AB' + str(row_number)).value.strip()
	# What size serving?
	serving_size = worksheet.acell('AC' + str(row_number)).value.strip()
	# Does your dog have any allergies or digestive problems?
	allergies_digestive = worksheet.acell('AD' + str(row_number)).value.strip()
	# Is your dog allowed to have treats?
	treats_ok = worksheet.acell('AE' + str(row_number)).value.strip()
	# Should your dog be fed during daycare?
	fed_during_daycare = worksheet.acell('AF' + str(row_number)).value.strip()
	# IN YOUR HOME SECTION
	# Is your dog restricted to a certain area of your home?
	home_restriction = worksheet.acell('AG' + str(row_number)).value.strip()
	# Do you leave water out all the time?
	water_out = worksheet.acell('AH' + str(row_number)).value.strip()
	# Dry food?
	dry_food = worksheet.acell('AI' + str(row_number)).value.strip()
	# Where do you keep the leash and collar/harness?
	leash_location = worksheet.acell('AJ' + str(row_number)).value.strip()
	# Where do you keep dog food, treats, and feeding/water dishes?
	where_is_the_stuff = worksheet.acell('AK' + str(row_number)).value.strip()
	# Where do you keep other items we might need (i.e. wee-wee pads, paper towels, toys)?
	where_is_the_other_stuff = worksheet.acell('AL' + str(row_number)).value.strip()
	# Do you have any instructions regarding lights and locks?
	lights_locks = worksheet.acell('AM' + str(row_number)).value.strip()
	# Do you have any other helpful information such as hiding spots or personality quirks?
	helpful_info = worksheet.acell('AN' + str(row_number)).value.strip()
	# Please describe where your dog is allowed in your home.
	dog_allowed = worksheet.acell('AO' + str(row_number)).value.strip()
	# Please describe your dog's behavior around other dogs.
	behavior = worksheet.acell('AP' + str(row_number)).value.strip()
	# Please describe your dog's behavior around new people:
	behavior_new_people = worksheet.acell('AQ' + str(row_number)).value.strip()
	# Please describe your dog's behavior on a leash.
	behavior_leash = worksheet.acell('AR' + str(row_number)).value.strip()
	# Is your dog housebroken?
	housebroken = worksheet.acell('AS' + str(row_number)).value.strip()
	# When was you last visit to the vet?
	last_vet = worksheet.acell('AT' + str(row_number)).value.strip()
	# What type of flea prevention do you use?
	flea_prevention = worksheet.acell('AU' + str(row_number)).value.strip()
	# Does your dog have any medical conditions?
	medical_conditions = worksheet.acell('AV' + str(row_number)).value.strip()
	# Does your dog require any medications?
	medications = worksheet.acell('AW' + str(row_number)).value.strip()
	# Has your dog ever bitten or tried to attack another dog or human being? (If yes, please elaborate below.) 
	ever_attacked = worksheet.acell('AX' + str(row_number)).value.strip()
	# Please list your dog's medications. 
	medication_list = worksheet.acell('AY' + str(row_number)).value.strip()
	# Please describe your dog's medication conditions.
	describe_medical_conditions =  worksheet.acell('AZ' + str(row_number)).value.strip()
	#VETERINARIAN SECTION
	vet_name = worksheet.acell('BA' + str(row_number)).value.strip()
	vet_phone = worksheet.acell('BB' + str(row_number)).value.strip()
	vet_address = worksheet.acell('BC' + str(row_number)).value.strip()
	vet_city = worksheet.acell('BD' + str(row_number)).value.strip()
	vet_state = worksheet.acell('BE' + str(row_number)).value.strip()
	# Do you agree with the terms?
	terms_agreement = worksheet.acell('BF' + str(row_number)).value.strip()
	# Receives another email address here for some reason. 
	email_again = worksheet.acell('BG' + str(row_number)).value.strip()
	
	'''
	##### CHECK FOR EMPTY FIELDS #####
	#### Check for empty fields in registration. Maybe picked the wrong row?
	if last_name or first_name or pet_name == '':
		blank_ok = input("\nAwww man, it looks like some fields are blank and you might get an empty document.\n\tAre you sure you still want to print row " + str(row_number) + "?  (y or n)\n\t")
		if blank_ok == 'y':
			print("\n\tOkay! I'll send you what I have ...")
			c_doc.create_docx(row_number)
			# Do shit
		elif blank_ok == 'n':
			print("\n\tNo problem! Heading back to the home screen ...")
			
			# Do other shit	
		else:
			print("\nUnrecognized response. Please type y or n.\n")
	else:
		pass
	'''