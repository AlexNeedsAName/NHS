#!/usr/bin/env python3
import gspread
import csv
import httplib2
import googleapiclient.discovery
import googleapiclient.http
from oauth2client.service_account import ServiceAccountCredentials
from datetime import date

KEYS = ['A', 'E', 'P']

def round(number, nearest):
	return number // nearest * nearest

def fillRow(sheet, row, data):
	for col in range(sheet.col_count):
		col+=1
		key = sheet.cell(1,col).value
		try:
			sheet.update_cell(row, col, data[key])
		except KeyError:
			sheet.update_cell(row, col, None)

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
today = date.today()
offset = today.year - int(str(today.year)[-2:])
for meeting in meetings:
	try:
		month, day, year = [ int(s) for s in meeting.split('-') ]
		year += offset
		meeting_date = date(year, month, day)
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

# Update the sheet
for i,person in enumerate(people.values()):
	i+=2
	print(person["Name"])
	for key in KEYS:
		person[key] = list(person.values()).count(key)
	fillRow(overview, i, person)
