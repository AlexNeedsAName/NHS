#!/usr/bin/env python3
import json
import gspread
import sys
import httplib2
import googleapiclient.discovery
import googleapiclient.http
import time
import csv
from oauth2client.service_account import ServiceAccountCredentials
from collections import namedtuple

Entry = namedtuple('Entry', ['date', 'task', 'hours', 'contact', 'photo'])

REQUIRED_IN_HOURS = 10
REQUIRED_OUT_HOURS = 10

USER_SHEET_TITLE = "{full_name}'s Hours"
USER_WELCOME_MESSAGE = "Hi {first_name}, this is the spreadsheet you can use to view your logged hours. Please save this to your school account's Google Drive. Please allow for up to 24 hours for new activities to appear."
USER_SHEET_FIRST_ROW = 3

def readConfig():
	with open('config.json', 'r') as f:
		CONFIG = json.load(f)
	return CONFIG

def writeConfig(CONFIG):
	with open('config.json', 'w') as f:
		f.write(json.dumps(CONFIG, indent=4))

CONFIG = readConfig()

class Person:
	def __init__(self, email):
		self.email = email
		self.in_hours = Hours(REQUIRED_IN_HOURS)
		self.out_hours = Hours(REQUIRED_OUT_HOURS)
		try:
			self.name = names[email]
		except KeyError:
			self.name = email

	def addHours(self, row):
		if(row['Type of Hours'] == 'In Hours'):
			hours = self.in_hours
		else:
			hours = self.out_hours
		hours.addEntry(row)

class Hours:
	def __init__(self, required_hours):
		self.total = 0
		self.required = required_hours
		self.entries = []

	def addEntry(self, row):
		entry = Entry(
			row['Date of Service'],
			row['Task/Type of Service'],
			row['Number of Service Hours'],
			row['Contact of Service Supervisor'],
			row['Photo of Signed Hour Sheet'],
		)
		self.entries.append(entry)
		self.total += entry.hours

	def update(self,worksheet):
		rows = len(self.entries)
		for i in range(rows):
			row_index = i+2
			row = worksheet.range('A{0}:E{0}'.format(row_index))
			for cell,value in zip(row,self.entries[i]):
				cell.value = value
			worksheet.update_cells(row)

	def getEntries(self):
		return self.entries

	def getTotal(self):
		return self.total

	def getRemaining(self):
		remaining = self.required - self.total
		if(remaining < 0):
			remaining = 0
		return remaining

def updateOverview(person, worksheet, i, hide_detail=False):
	person_data = [ person.email,
					person.name,
	                person.in_hours.getTotal(),
	                person.in_hours.getRemaining(),
	                person.out_hours.getTotal(),
	                person.out_hours.getRemaining(),
	                "https://docs.google.com/spreadsheets/d/{id}".format(id=sheet.id),
	]

	if(hide_detail):
		row = worksheet.range('A{0}:F{0}'.format(i))
	else:
		row = worksheet.range('A{0}:G{0}'.format(i))
	for cell,value in zip(row,person_data):
		cell.value = value
	worksheet.update_cells(row)

def createNewSheet(person):
	id = client.open("Template").id
	copied_file = {"title": USER_SHEET_TITLE.format(full_name=person.name)}
	drive_service.files().copy(fileId=id, body=copied_file).execute()
	print("Created new sheet for {}".format(person.name))

def shareNewSheet(person):
	sheet = client.open(USER_SHEET_TITLE.format(full_name=person.name))
	sheet.share(person.email, perm_type='user', role='reader', email_message=USER_WELCOME_MESSAGE.format(first_name=person.name.split(' ')[0]))
	for email in CONFIG["HOURS"]["ADMIN_EMAILS"]:
		sheet.share(email, perm_type='user', role='reader', notify=False)
	print("Shared new sheet for {}".format(person.name))
	return sheet


def process():
	# use creds to create a client to interact with the Google Drive API
	scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
	creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
	client = gspread.authorize(creds)

	# Create an instance of google drive service
	http = httplib2.Http()
	creds.authorize(http)
	drive_service = googleapiclient.discovery.build('drive', 'v2', http=http)

	# Find a Sheet by name and open the first sheet
	spreadsheet = client.open("SLEHS NHS Hour Submission (Responses)")
	responses = spreadsheet.worksheet("Responses")
	overview = spreadsheet.worksheet("Overview")

	# read the database of people and their emails
	with open('people.csv', mode='r') as file:
		reader = csv.reader(file)
		names = dict(reader)

	# Extract the emails of each person
	people = {}
	data = responses.get_all_records()

	if(len(data) <= CONFIG["HOURS"]["LAST_CHECKED_ENTRIES"]):
		print("No new entries")
		sys.exit(0)

	for row in data:
		email = row["Email Address"]
		if(email not in people.keys()):
			people[email] = Person(email)
		people[email].addHours(row)
	print("Parsed all {} entries for {} users".format(len(data), len(people)))

	#We're done with the dict of all emails and names now. We have the ones we need.
	del names

	# Create, update, and share individual sheets for each user.
	for i,email in enumerate(people):
		i+=USER_SHEET_FIRST_ROW
		person = people[email]

		#Get user's sheet
		try:
			sheet = client.open(USER_SHEET_TITLE.format(full_name=person.name))
		except gspread.exceptions.SpreadsheetNotFound:
			createNewSheet(person)
			sheet = shareNewSheet(person)

		#Update In Hours
		in_hours = sheet.worksheet("In Hours")
		person.in_hours.update(in_hours)

		#Update Out Hours
		out_hours = sheet.worksheet("Out Hours")
		person.out_hours.update(out_hours)

		#Update personal overview and the main overview
		personal_overview = sheet.worksheet("Overview")
		updateOverview(person, personal_overview, 3, hide_detail=True)
		updateOverview(person, overview, i)

		print("Done updating {}'s hours".format(person.name))

	CONFIG["HOURS"]["LAST_CHECKED_ENTRIES"] = len(data)
	writeConfig(CONFIG)

if(__name__ == "__main__"):
	process()
