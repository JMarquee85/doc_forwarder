#import main_loop
## pc_doc_forwarder Registration Function
# Josh Marcus
#from main_loop import gc

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


def get_reg():
	scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
	credentials = ServiceAccountCredentials.from_json_keyfile_name(json_key, scope)
	gc = gspread.authorize(credentials)
	
	# Pulling Today's Date
	now = datetime.now()
	today_date = now.strftime("%m/%d/%Y %I:%M")
	
	# docx Template File (pc Registration Form with MailMerge Fields)
	template = os.path.join(this_dir, 'pcregtemplate.docx')
	# Create a mailmerge document
	document = MailMerge(template)	
	
	print("\n\nSCANNING REGISTRATION DATABASE FOR NEW APPLICATIONS ...\n\t\t(ctrl-c to quit)")
	reg = gc.open_by_url('https://docs.google.com/spreadsheets/d/1dYO0M9iBWVmOYcE8fO9t9rzIaTijrbLQBCNYbulAuF4/edit#gid=136975089')
	worksheet = reg.get_worksheet(0)
	### Variable Names ###
	###### Consider making the columns here contextual and set by
	###### argument set in this function. Accept row number by user
	###### input in console. 
	# Owner Information Section
	# Assigning variables to first row response in Google Sheet
	time_registered = worksheet.acell('A2').value.strip()
	last_name = worksheet.acell('B2').value.strip()
	first_name = worksheet.acell('C2').value.strip()
	street = worksheet.acell('D2').value.strip()
	apt = worksheet.acell('E2').value.strip()
	cross_streets = worksheet.acell('F2').value.strip()
	city = worksheet.acell('G2').value.strip()
	state = worksheet.acell('H2').value.strip()
	zip_code = worksheet.acell('I2').value.strip()
	home_phone = worksheet.acell('J2').value.strip()
	cell_phone = worksheet.acell('K2').value.strip()
	business_phone = worksheet.acell('L2').value.strip()
	email_address = worksheet.acell('M2').value.strip()
	# Where did you hear about us?
	reference = worksheet.acell('N2').value.strip()
	emergency_contacts = worksheet.acell('O2').value.strip()
	membership = worksheet.acell('P2').value.strip()
	keys = worksheet.acell('Q2').value.strip()
	# Pet Information Section
	pet_name = worksheet.acell('R2').value.strip()
	pet_nick = worksheet.acell('S2').value.strip()
	breed = worksheet.acell('T2').value.strip()
	weight = worksheet.acell('U2').value.strip()
	sex = worksheet.acell('V2').value.strip()
	dob = worksheet.acell('W2').value.strip()
	color = worksheet.acell('X2').value.strip()
	fixed = worksheet.acell('Y2').value.strip()
	# How long have you owned your dog?
	length_owned = worksheet.acell('Z2').value.strip()
	# Please specify what brand and type (wet/ dry) food you use:
	brand_food = worksheet.acell('AA2').value.strip()
	# How many times per day do you feed your dog?
	times_per_day_feed = worksheet.acell('AB2').value.strip()
	# What size serving?
	serving_size = worksheet.acell('AC2').value.strip()
	# Does your dog have any allergies or digestive problems?
	allergies_digestive = worksheet.acell('AD2').value.strip()
	# Is your dog allowed to have treats?
	treats_ok = worksheet.acell('AE2').value.strip()
	# Should your dog be fed during daycare?
	fed_during_daycare = worksheet.acell('AF2').value.strip()
	# IN YOUR HOME SECTION
	# Is your dog restricted to a certain area of your home?
	home_restriction = worksheet.acell('AG2').value.strip()
	# Do you leave water out all the time?
	water_out = worksheet.acell('AH2').value.strip()
	# Dry food?
	dry_food = worksheet.acell('AI2').value.strip()
	# Where do you keep the leash and collar/harness?
	leash_location = worksheet.acell('AJ2').value.strip()
	# Where do you keep dog food, treats, and feeding/water dishes?
	where_is_the_stuff = worksheet.acell('AK2').value.strip()
	# Where do you keep other items we might need (i.e. wee-wee pads, paper towels, toys)?
	where_is_the_other_stuff = worksheet.acell('AL2').value.strip()
	# Do you have any instructions regarding lights and locks?
	lights_locks = worksheet.acell('AM2').value.strip()
	# Do you have any other helpful information such as hiding spots or personality quirks?
	helpful_info = worksheet.acell('AN2').value.strip()
	# Please describe where your dog is allowed in your home.
	dog_allowed = worksheet.acell('AO2').value.strip()
	# Please describe your dog's behavior around other dogs.
	behavior = worksheet.acell('AP2').value.strip()
	# Please describe your dog's behavior around new people:
	behavior_new_people = worksheet.acell('AQ2').value.strip()
	# Please describe your dog's behavior on a leash.
	behavior_leash = worksheet.acell('AR2').value.strip()
	# Is your dog housebroken?
	housebroken = worksheet.acell('AS2').value.strip()
	# When was you last visit to the vet?
	last_vet = worksheet.acell('AT2').value.strip()
	# What type of flea prevention do you use?
	flea_prevention = worksheet.acell('AU2').value.strip()
	# Does your dog have any medical conditions?
	medical_conditions = worksheet.acell('AV2').value.strip()
	# Does your dog require any medications?
	medications = worksheet.acell('AW2').value.strip()
	# Has your dog ever bitten or tried to attack another dog or human being? (If yes, please elaborate below.) 
	ever_attacked = worksheet.acell('AX2').value.strip()
	# Please list your dog's medications. 
	medication_list = worksheet.acell('AY2').value.strip()
	# Please describe your dog's medication conditions.
	describe_medical_conditions =  worksheet.acell('AZ2').value.strip()
	#VETERINARIAN SECTION
	vet_name = worksheet.acell('BA2').value.strip()
	vet_phone = worksheet.acell('BB2').value.strip()
	vet_address = worksheet.acell('BC2').value.strip()
	vet_city = worksheet.acell('BD2').value.strip()
	vet_state = worksheet.acell('BE2').value.strip()
	# Do you agree with the terms?
	terms_agreement = worksheet.acell('BF2').value.strip()
	# Receives another email address here for some reason. 
	email_again = worksheet.acell('BG2').value.strip()
	
	print("\nChecking database for last detected registration: \n\t" + pet_name.title().strip() + " " + last_name.title().strip() + "\n\tOwner: " +
				"\t" + last_name.title().strip() + ", " + first_name.title().strip())

	##### THIS FUNCTION CHECKS IF CUSTOMER HAS ALREADY SUBMITTED 
	test_strings = (last_name, first_name, pet_name)
	
	### There's an issue here ... Need to create a list to check against
	### for recent applications 
	
	if last_name.title().strip() and first_name.title().strip() and pet_name.title().strip() in open('inputlog.txt').read():
	#if any(s in line for s in test_strings):
		print("\nThis customer's registration form has already been uploaded!")
		return
	else:
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
		file_to_convert = last_name.title().strip() + ', ' + pet_name.title().strip() + '.docx'
		converted_pdf = last_name.title().strip() + ', ' + pet_name.title().strip() + '.pdf'
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
		
		# RUN EMAIL FUNCTION
		e_pc.email_pupculture()
		
		return


