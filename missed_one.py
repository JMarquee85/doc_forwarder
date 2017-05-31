# This will be a function to accept a number as an argument
# that will represent the row that will be pulled
# as a temporary stopgap to grab documents the scripy may
# have missed

# Look into option to select from a list
# with that ncurses library I ran accross

from __future__ import print_function
from requests.exceptions import ConnectionError

##### MAIN IMPORT STATEMENTS #####
from mailmerge import MailMerge
from datetime import datetime
import dropbox
# See https://pypi.python.org/pypi/dropbox for more information

##### MY IMPORTS #####
### Private configuration file
import private_config as p_con
### Email PupCulture Function
import email_pupculture as e_pc
### Get Registration Function
from get_reg import get_reg

import gspread
import json
import sys
import os
import smtplib
import csv
import socket
import urllib, re
from oauth2client.service_account import ServiceAccountCredentials

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
json_key = os.path.join(this_dir, 'pcregisterpdfforwarder-1f39ce9a4ac0.json')

##### SCOPE #####
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
# Google Credentials
credentials = ServiceAccountCredentials.from_json_keyfile_name(json_key, scope)
global gc
gc = gspread.authorize(credentials)

# Open Registration Google Spreadsheet
reg = gc.open_by_url('https://docs.google.com/spreadsheets/d/1dYO0M9iBWVmOYcE8fO9t9rzIaTijrbLQBCNYbulAuF4/edit#gid=136975089')
worksheet = reg.get_worksheet(0)
print("Google sheet successfully opened!")

##### 
# For more information on using MailMerge:
# pbpython.com/python-word-template.html
# Using templates with docx-mailmerge

# Template File
template = os.path.join(this_dir, 'pcregtemplate.docx')

#####  THE FUNCTION #####
###
#

def reprint_row(row_select):
	# Check to make sure the input represents
	# a row with customer content in it
	if row_select <= 1:
		print("\n\nPlease input a number of 2 or higher! (Note: 2 is the first row in the Google Doc.)")
		return
	else:
		print("\nImporting customer information from Row " + 
				str(str(row_select)) + "...")
		# Input rowselect number to call particular row
		time_registered = worksheet.acell('A' + str(str(row_select))).value
		last_name = worksheet.acell('B' + str(str(row_select))).value
		first_name = worksheet.acell('C' + str(str(row_select))).value
		street = worksheet.acell('D' + str(row_select)).value
		apt = worksheet.acell('E' + str(str(row_select))).value
		cross_streets = worksheet.acell('F' + str(row_select)).value
		city = worksheet.acell('G' + str(row_select)).value
		state = worksheet.acell('H' + str(row_select)).value
		zip_code = worksheet.acell('I' + str(row_select)).value
		home_phone = worksheet.acell('J' + str(row_select)).value
		cell_phone = worksheet.acell('K' + str(row_select)).value
		business_phone = worksheet.acell('L' + str(row_select)).value
		email_address = worksheet.acell('M' + str(row_select)).value
		# Where did you hear about us?
		reference = worksheet.acell('N' + str(row_select)).value
		emergency_contacts = worksheet.acell('O' + str(row_select)).value
		membership = worksheet.acell('P' + str(row_select)).value
		keys = worksheet.acell('Q' + str(row_select)).value
		# Pet Information Section
		pet_name = worksheet.acell('R' + str(row_select)).value
		pet_nick = worksheet.acell('S' + str(row_select)).value
		breed = worksheet.acell('T' + str(row_select)).value
		weight = worksheet.acell('U' + str(row_select)).value
		sex = worksheet.acell('V' + str(row_select)).value
		dob = worksheet.acell('W' + str(row_select)).value
		color = worksheet.acell('X' + str(row_select)).value
		fixed = worksheet.acell('Y' + str(row_select)).value
		# How long have you owned your dog?
		length_owned = worksheet.acell('Z' + str(row_select)).value
		# Please specify what brand and type (wet/ dry) food you use:
		brand_food = worksheet.acell('AA' + str(row_select)).value
		# How many times per day do you feed your dog?
		times_per_day_feed = worksheet.acell('AB' + str(row_select)).value
		# What size serving?
		serving_size = worksheet.acell('AC' + str(row_select)).value
		# Does your dog have any allergies or digestive problems?
		allergies_digestive = worksheet.acell('AD' + str(row_select)).value
		# Is your dog allowed to have treats?
		treats_ok = worksheet.acell('AE' + str(row_select)).value
		# Should your dog be fed during daycare?
		fed_during_daycare = worksheet.acell('AF' + str(row_select)).value
		# IN YOUR HOME SECTION
		# Is your dog restricted to a certain area of your home?
		home_restriction = worksheet.acell('AG' + str(row_select)).value
		# Do you leave water out all the time?
		water_out = worksheet.acell('AH' + str(row_select)).value
		# Dry food?
		dry_food = worksheet.acell('AI' + str(row_select)).value
		# Where do you keep the leash and collar/harness?
		leash_location = worksheet.acell('AJ' + str(row_select)).value
		# Where do you keep dog food, treats, and feeding/water dishes?
		where_is_the_stuff = worksheet.acell('AK' + str(row_select)).value
		# Where do you keep other items we might need (i.e. wee-wee pads, paper towels, toys)?
		where_is_the_other_stuff = worksheet.acell('AL' + str(row_select)).value
		# Do you have any instructions regarding lights and locks?
		lights_locks = worksheet.acell('AM' + str(row_select)).value
		# Do you have any other helpful information such as hiding spots or personality quirks?
		helpful_info = worksheet.acell('AN' + str(row_select)).value
		# Please describe where your dog is allowed in your home.
		dog_allowed = worksheet.acell('AO' + str(row_select)).value
		# Please describe your dog's behavior around other dogs.
		behavior = worksheet.acell('AP' + str(row_select)).value
		# Please describe your dog's behavior around new people:
		behavior_new_people = worksheet.acell('AQ' + str(row_select)).value
		# Please describe your dog's behavior on a leash.
		behavior_leash = worksheet.acell('AR' + str(row_select)).value
		# Is your dog housebroken?
		housebroken = worksheet.acell('AS' + str(row_select)).value
		# When was you last visit to the vet?
		last_vet = worksheet.acell('AT' + str(row_select)).value
		# What type of flea prevention do you use?
		flea_prevention = worksheet.acell('AU' + str(row_select)).value
		# Does your dog have any medical conditions?
		medical_conditions = worksheet.acell('AV' + str(row_select)).value
		# Does your dog require any medications?
		medications = worksheet.acell('AW' + str(row_select)).value
		# Has your dog ever bitten or tried to attack another dog or human being? (If yes, please elaborate below.) 
		ever_attacked = worksheet.acell('AX' + str(row_select)).value
		# Please list your dog's medications. 
		medication_list = worksheet.acell('AY' + str(row_select)).value
		# Please describe your dog's medication conditions.
		describe_medical_conditions =  worksheet.acell('AZ' + str(row_select)).value
		#VETERINARIAN SECTION
		vet_name = worksheet.acell('BA' + str(row_select)).value
		vet_phone = worksheet.acell('BB' + str(row_select)).value
		vet_address = worksheet.acell('BC' + str(row_select)).value
		vet_city = worksheet.acell('BD' + str(row_select)).value
		vet_state = worksheet.acell('BE' + str(row_select)).value
		# Do you agree with the terms?
		terms_agreement = worksheet.acell('BF' + str(row_select)).value
		# Receives another email address here for some reason. 
		email_again = worksheet.acell('BG' + str(row_select)).value
		
		# Variables changed to selected row. 
		# Printing confirmation message. 
		print("Creating docx file for "  + last_name.upper() + " " +
					pet_name.upper() + "...")
		# Create docx file			
					
		print("\nEmailing docx file!\n")
		# Email file to pupculture
		e_it.email_it()	
		
		return


##### DO THE BIT! #####

while True:
	# User inputs number to select row to reprint
	try:
		# Accepts user input by row number
		row_select = input("Please input the row number you would like to reprint and resend ... \n\t")
		reprint_row(int(row_select))
	except KeyboardInterrupt:
		print("\nOK! Exiting program!")
		break
	except ConnectionError:
		print("\nUnable to connect! Please ensure you are connected to the internet! \nTrying again!")
		slack.chat.post_message(p_con.slack_channel, 'Connection lost! Attempting to reconnect ... ')
		
	

	
