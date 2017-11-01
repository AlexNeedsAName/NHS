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
import argparse

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

class MyParser(argparse.ArgumentParser):
	def error(self, message):
		print("error: " + message)
		self.print_help()
		sys.exit(2)

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

def manual(state='P'):
	while(True):
		valid = False
		while(not valid):
			email = input("Email: ").lower()
			valid = (email in (key.lower() for key in names.keys()))
			if(not valid):
				print("Invalid school email.")
		mark(email, state, datetime.date.today())

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
	print("Reading Responses...")
	# Login to  google sheets
	scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
	creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
	client = gspread.authorize(creds)

	# Find a Sheet by name and open the first sheet
	spreadsheet = client.open("SLEHS NHS Attendance (Responses)")
	responses = spreadsheet.worksheet("Responses")
	overview = spreadsheet.worksheet("Overview")

	#Put the read functions first and together because they take a long time.
	data = responses.get_all_records()
	meetings = dict((date,'') for date in overview.row_values(1)[1:])

	print("Processing Data...")

	#First get the meetings and have the ones in the past default to Absent
	today = datetime.date.today()
	for meeting in meetings:
		try:
			month, day, year = [ int(s) for s in meeting.split('/') ]
			meeting_date = datetime.date(year, month, day)
			if(meeting_date < today):
				meetings[meeting] = 'A'
		except ValueError: #Not a meeting, some other header
			pass

	# Create an dict for each person and fill it's data
	people = {}

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

	last_cell = gspread.utils.rowcol_to_a1(overview.row_count,overview.col_count)
	cell_list = overview.range('A2:{}'.format(last_cell))

	rows = [cell_list[x:x+overview.col_count] for x in range(0, len(cell_list), overview.col_count)]

	FIRST_ROW = overview.row_values(1)

	# Update the sheet
	people = sorted(people.values(), key=lambda k: k['Name'])
	for r,person in enumerate(people):
		for key in KEYS:
			person[key] = list(person.values()).count(key)
		for c,key in enumerate(FIRST_ROW):
			rows[r][c].value = person[key]
	overview.update_cells(cell_list)

if(__name__ == "__main__"):
	parser = MyParser(description='A script for mannaging the attendance of an orginization or a class.')
	parser.add_argument("-m", "--manual", action='store_true', help="Take manual attendance.")
	parser.add_argument("-e", "--excuse", action='store_true', help="Take manual attendance.")
	parser.add_argument("-t", "--take",   action='store_true', help="Take attendance normally")
	parser.add_argument("-u", "--update", action='store_true', help="Update older entries")
	args = parser.parse_args()

	try:
		if(args.update):
			updateOldEntries()
		elif(args.manual):
			manual()
		elif(args.excuse):
			manual(state='E')
		elif(args.take):
			takeAttendance()
	except KeyboardInterrupt:
		print("\n")
	process()
	print("Done!")

