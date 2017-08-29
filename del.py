#!/usr/bin/env python3
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from collections import namedtuple


# use creds to create a client to interact with the Google Drive API
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)

print(client.openall())


print("Enter IDs:")
while(1):
	id = input()
	client.del_spreadsheet(id)
	print("Deleted.")
