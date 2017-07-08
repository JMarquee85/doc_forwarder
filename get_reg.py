#import main_loop
## pc_doc_forwarder Registration Function
# Josh Marcus
#from main_loop import gc

import gspread
import os
from oauth2client.service_account import ServiceAccountCredentials
from mailmerge import MailMerge
from datetime import datetime

##### Need to finish this off with two main things:
## 1. Need a way to make sure that no registrations were missed
##		if the program is not running all the time
## 2. Need to find a way to make sure this runs all the time
		# Run on pc windows machine? Server? Nothing always on.

## If time, or for later, it would be great to come up with a
## way to search by last name, dog name, etc. and reprint 
## rather than a row number

## Try to include that ncurses list

## Eventually push this information into a webapp?
# Pulling Today's Date
now = datetime.now()
today_date = now.strftime("%m/%d/%Y %I:%M")
# from slackclient import SlackClient
from slacker import Slacker

import private_config as p_con
import email_pupculture as e_pc
import import_gsheet_by_row as ig
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


##### CHECK TO SEE IF CUSTOMER FILE HAS ALREADY BEEN LOGGED #####
def check_for_file(row_number):
	
	last_name = worksheet.acell('B' + str(row_number)).value.strip()
	first_name = worksheet.acell('C' + str(row_number)).value.strip()
	pet_name = worksheet.acell('R' + str(row_number)).value.strip()

	if last_name.title().strip() and first_name.title().strip() and pet_name.title().strip() in open('inputlog.txt').read():
	#if any(s in line for s in test_strings):
		print("\nRegistration for " + pet_name.title().strip() + " " + last_name.title().strip() + " found!")
		return
	else:
		ig.import_gsheet_by_row(row_number)
		create_docx(row_number)
		e_it.email_it(row_number)
		return
		

##### This for loop checks the top twenty entries
##### to see if they have already registered. 		
def check_top_twenty():
	for x in range(2,20):
		print("\n\nCHECKING IF ENTRY " + str(x) + " has registered ...\n\t\t(ctrl-c to quit)")
		ig.import_gsheet_by_row(x)
		check_for_file(x)

##### CREATE DOCX FILE #####
def create_docx(row_number):
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

	
	#last_name = worksheet.acell('B' + str(row_number)).value.strip()
	#first_name = worksheet.acell('C' + str(row_number)).value.strip()
	#pet_name = worksheet.acell('R' + str(row_number)).value.strip()
	
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
	'''
	### Convert docx to pdf and store in folder called pdfs
	# Call w2p here
	file_to_convert = last_name.title() + ', ' + pet_name.title() + '.docx'
	converted_pdf = last_name.title() + ', ' + pet_name.title() + '.pdf'
	w2p.convx_to_pdf(file_to_convert, converted_pdf)
	'''
	### Post Slack Message to Update Channel
	slack.chat.post_message(p_con.slack_channel, new_reg_msg)
	
	##### LOGGING TO TEXT FILE #####
	with open('inputlog.txt', 'a') as doglog:
		doglog.write('\n' + last_name.title().strip() + ', ' + first_name.title().strip() + ' (Owner), ' + pet_name.title().strip() + ':  ' + str(today_date)) 
	### Confirm Logging ###
	print("Registration information for " + pet_name.title().strip() + " " + last_name.title().strip() + 
				" has been logged to inputlog.txt!")
		
	return


