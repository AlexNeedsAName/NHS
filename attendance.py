#!/usr/bin/env python3
import gspread
import csv
import httplib2
import json
import requests
import googleapiclient.discovery
import googleapiclient.http
from oauth2client.service_account import ServiceAccountCredentials
import datetime

KEYS = ['A', 'E', 'P']
FORM_URL = "https://docs.google.com/forms/d/e/{id}/formResponse?{data}"

def readConfig():
	with open('config.json', 'r') as f:
		CONFIG = json.load(f)
	return CONFIG

def writeConfig(CONFIG):
	with open('config.json', 'w') as f:
		f.write(json.dumps(CONFIG, indent=4))

CONFIG = readConfig()

def round(number, nearest):
	return number // nearest * nearest

def fillRow(sheet, row, data, headers=None):
	for col in range(1,sheet.col_count+1):
		key = sheet.cell(1,col).value
		try:
			sheet.update_cell(row, col, data[key])
		except KeyError:
			sheet.update_cell(row, col, "")

def submitForm(form_id, data):
	data_string = ''
	for i,keyvalue in enumerate(data.items()):
		key,value = keyvalue
		data_string += "{}={}".format(key,value)
		if(i+1 != len(data)):
			data_string += '&'
	url = FORM_URL.format(id=form_id, data=data_string)
	print(url)
	response = requests.get(url)
	print(response)
	return response

def mark(person, state, date):
	data = {
		CONFIG["ATTENDANCE"]["EMAIL"]: person,
		CONFIG["ATTENDANCE"]["STATE"]: state,
		CONFIG["ATTENDANCE"]["DATE"]:  "{year}-{month:0{width}}-{day:0{width}}".format(year=date.year, month=date.month, day=date.day, width=2),
	}
	submitForm(CONFIG["ATTENDANCE"]["FORM_ID"], data)

def takeAttendance():
	

def process():
	# Login to  google sheets
	scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
	creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
	client = gspread.authorize(creds)

	# Find a Sheet by name and open the first sheet
	spreadsheet = client.open("SLEHS NHS Attendance (Responses)")
	responses = spreadsheet.worksheet("Responses")
	overview = spreadsheet.worksheet("Overview")

	# Read the database of people and their emails
	with open('people.csv', mode='r') as file:
		reader = csv.reader(file)
		names = dict(reader)

	#First get the meetings and have the ones in the past default to Absent
	meetings = dict((date,'') for date in overview.row_values(1)[1:])
	today = datetime.date.today()
	offset = today.year - int(str(today.year)[-2:])
	for meeting in meetings:
		try:
			month, day, year = [ int(s) for s in meeting.split('-') ]
			year += offset
			meeting_date = datetime.date(year, month, day)
			if(meeting_date <= today):
				meetings[meeting] = 'A'
		except ValueError: #Not a meeting, some other header
			pass

	# Create an object for each person and fill it's data
	people = {}
	data = responses.get_all_records()

	for row in data:
		email = row["Student Email"]
		date = row["Date"]
		if(email not in people):
			people[email] = meetings.copy()
			people[email]["Name"] = names[email]
		if(date in meetings):
			people[email][date] = row["State"]

	# We're done with the dict of all emails and names now. We have the ones we need.
	del names

	#Resize the spreadsheet
	if(len(people) > 0):
		overview.resize(rows=len(people)+1)
	else:
		overview.resize(rows=2)
		fillRow(overview, 2, {})

	# Update the sheet
	for i,person in enumerate(people.values()):
		i+=2
		print(person["Name"])
		for key in KEYS:
			person[key] = list(person.values()).count(key)
		fillRow(overview, i, person)

if(__name__ == "__main__"):
#	mark("martiale002@slyon.us", "P", datetime.date.today())
	process()
