#!/usr/bin/env python3
import time
import gspread
import csv
import sys
import httplib2
import json
import requests
import googleapiclient.discovery
import googleapiclient.http
from oauth2client.service_account import ServiceAccountCredentials
import datetime
import serial

START = b'\x02'
END =   b'\x03'

KEYS = ['A', 'E', 'P']
FORM_URL = "https://docs.google.com/forms/d/e/{id}/formResponse?{data}"

with open('people.csv', mode='r') as file:
	reader = csv.reader(file)
	names = dict(reader)

with open('ids.csv', mode='r') as file:
	reader = csv.reader(file)
	emails = dict(reader)

def readConfig():
	with open('config.json', 'r') as f:
		CONFIG = json.load(f)
	return CONFIG

CONFIG = readConfig()

class scanner:
	def __init__(self, serial_port="/dev/cu.SLAB_USBtoUART", debounce_delay=.2):
		self.s = serial.Serial(serial_port, timeout=debounce_delay)

	def readID(self):
		id = b''
		while(True):
			c = self.s.read()
			if(c == START):
				id = b''
			elif(c == END):
				id = id.decode("utf-8")
				break
			else:
				id += c
		print(id)
		while(len(self.s.read()) > 0):
			pass
		return id

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
	#print(url)
	response = requests.get(url)
	#print(response)
	return response.status_code

def mark(person, state, date):
	data = {
		CONFIG["ATTENDANCE"]["EMAIL"]: person,
		CONFIG["ATTENDANCE"]["STATE"]: state,
		CONFIG["ATTENDANCE"]["DATE"]:  "{year}-{month:0{width}}-{day:0{width}}".format(year=date.year, month=date.month, day=date.day, width=2),
	}
	status = submitForm(CONFIG["ATTENDANCE"]["FORM_ID"], data)
	if(status != 200):
		print("Could not connect to google. Saving offline.")
		with open("offline.csv") as file:
			file.write("{},{},{}".format(person, state, date))

def takeAttendance():
	global emails

	print("Connecting to serial... ", end='')
	sys.stdout.flush()
	scan = scanner()
	print("Connected.\nReady to take attendance.")

	while(True):
		print("ID:", end=' ')
		sys.stdout.flush()
		id = scan.readID()

		if(id not in emails.keys()):
			register(id)
		email = emails[id]
		name = names[email]
		print("Welcome, {}\n".format(name))
		mark(email, 'P', datetime.date.today())

def manual():
	while(True):
		valid = False
		while(not valid):
			email = input("Email: ")
			valid = (email.lower() in (key.lower() for key in names.keys()))
			if(not valid):
				print("Invalid school email.")
		mark(email, 'P', datetime.date.today())

def updateOldEntries():
	while(True):
		valid = False
		while(not valid):
			email = input("Email: ")
			valid = (email.lower() in (key.lower() for key in names.keys()))
			if(not valid):
				print("Invalid school email.")
		state = input("State: ").upper()
		date_string = input("Date: ")
		month, day, year = [ int(s) for s in date_string.split('/') ]
		date = datetime.date(year, month, day)
		mark(email, state, date)

def register(id):
	valid = False
	while(not valid):
		email = input("Enter you school email: ")
		valid = (email.lower() in (key.lower() for key in names.keys()))
		if(not valid):
			print("Invalid school email.")
	emails[id] = email
	with open('ids.csv', 'a') as f:
		f.write("{},{}\n".format(id,email))

def process():
	# Login to  google sheets
	scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
	creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
	client = gspread.authorize(creds)

	# Find a Sheet by name and open the first sheet
	spreadsheet = client.open("SLEHS NHS Attendance (Responses)")
	responses = spreadsheet.worksheet("Responses")
	overview = spreadsheet.worksheet("Overview")

	#First get the meetings and have the ones in the past default to Absent
	meetings = dict((date,'') for date in overview.row_values(1)[1:])
	today = datetime.date.today()
	for meeting in meetings:
		try:
			month, day, year = [ int(s) for s in meeting.split('/') ]
			meeting_date = datetime.date(year, month, day)
			print(meeting_date)
			if(meeting_date <= today):
				print(True)
				meetings[meeting] = 'A'
		except ValueError: #Not a meeting, some other header
			pass

	# Create an dict for each person and fill it's data
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

	#Resize the spreadsheet
	overview.resize(rows=2)
	if(len(people) > 0):
		overview.resize(rows=len(people)+1)
	else:
		overview.resize(rows=2)
		fillRow(overview, 2, {})

	print("Processed data. Updating spreadsheet...")

	# Update the sheet
	n = len(people)
	people = sorted(people.values(), key=lambda k: k['Name'])
	for i,person in enumerate(people):
		print("Updating {} of {}\r".format(i+1,n))
		i+=2
		for key in KEYS:
			person[key] = list(person.values()).count(key)
		fillRow(overview, i, person)

if(__name__ == "__main__"):
#	process()

	try:
		updateOldEntries()
#		manual()
#		takeAttendance()
	except KeyboardInterrupt:
		print("\nProcessing data...")
		process()
		print("Done!")

