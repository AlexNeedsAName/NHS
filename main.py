#!/usr/bin/env python3
import gspread
import sys
import httplib2
import googleapiclient.discovery
import googleapiclient.http
import time
from oauth2client.service_account import ServiceAccountCredentials
from collections import namedtuple

Entry = namedtuple('Entry', ['date', 'task', 'hours', 'contact', 'photo'])

REQUIRED_IN_HOURS = 10
REQUIRED_OUT_HOURS = 10

ADMIN_EMAILS = ["SCHMITZP@slcs.us", "martiale002@slyon.us", "churcsam000@slyon.us", "adamsbro000@slyon.us", "renehfal000@slyon.us", "aulicnat000@slyon.us", "pavlicia000@slyon.us"]

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

class Person:
	def __init__(self, email):
		self.email = email
		self.in_hours = Hours(REQUIRED_IN_HOURS)
		self.out_hours = Hours(REQUIRED_OUT_HOURS)

	def addHours(self, row):
		if(row['Type of Hours'] == 'In Hours'):
			hours = self.in_hours
		else:
			hours = self.out_hours
		hours.addEntry(row)

def updateOverview(person, worksheet, i, hide_detail=False):
	person_data = [person.email, person.in_hours.getTotal(), person.in_hours.getRemaining(), person.out_hours.getTotal(), person.out_hours.getRemaining(), "https://docs.google.com/spreadsheets/d/{id}".format(id=sheet.id)]
	if(hide_detail):
		row = worksheet.range('A{0}:E{0}'.format(i))
	else:
		row = worksheet.range('A{0}:F{0}'.format(i))
	for cell,value in zip(row,person_data):
		cell.value = value
	worksheet.update_cells(row)

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

# Extract the emails of each person
people = {}
data = responses.get_all_records()
for row in data:
	email = row["Email Address"]
	if(email not in people.keys()):
		people[email] = Person(email)
	people[email].addHours(row)
print("Parsed all {} entries for {} users".format(len(data), len(people)))

# Create, update, and share individual sheets for each user.
for i,person in enumerate(people):
	i+=3
	name = person.replace("@slyon.us", '')
	person = people[person]

	#Get detail sheet
	try:
		sheet = client.open("{}'s Hours".format(name))
	except gspread.exceptions.SpreadsheetNotFound:
		id = client.open("Template").id
		copied_file = {"title": "{}'s Hours".format(name)}
		drive_service.files().copy(fileId=id, body=copied_file).execute()
		sheet = client.open("{}'s Hours".format(name))
		sheet.share(person.email, perm_type='user', role='reader', email_message="Hi, this is the spreadsheet you can use to view your logged hours. It will be updated every day at midnight.")
		for email in ADMIN_EMAILS:
			sheet.share(email, perm_type='user', role='reader', notify=False)
		print("Created and shared new Sheet for {}".format(name))
	in_hours = sheet.worksheet("In Hours")
	out_hours = sheet.worksheet("Out Hours")
	personal_overview = sheet.worksheet("Overview")

	person.in_hours.update(in_hours)
	person.out_hours.update(out_hours)
	updateOverview(person, overview, i)
	updateOverview(person, personal_overview, 3, hide_detail=True)

	print("Done updating {}'s hours".format(name))
