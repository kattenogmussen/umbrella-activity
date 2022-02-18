import requests
import json
import os
import getpass
from pprint import pprint
from oauthlib.oauth2 import BackendApplicationClient
from oauthlib.oauth2 import TokenExpiredError
from requests_oauthlib import OAuth2Session
from requests.auth import HTTPBasicAuth
import urllib
import csv
import xlsxwriter
import time
import sys
from datetime import datetime
import ipaddress

print('This script takes a source input file consiting of IPs, DOMAINs or URLs and outputs a spreadsheet with a summary or all requestst registred in the Umbrella dashboard.')

def san_query(query):
    query = query.split('?')[0]
    query = query.split('#')[0]
    return urllib.parse.quote(query)

def check_if_umbrella_list(file):
    with open(file) as csv_file:
        csv_reader = csv.DictReader(csv_file)
        dict_from_csv = dict(list(csv_reader)[0])
        list_of_column_names = list(dict_from_csv.keys())
        list_of_umbrella_collumns = ["id","destination","type","comment","createdAt"]
        return all(elem in list_of_umbrella_collumns for elem in list_of_column_names)

def check_if_file_exists(filename):
    try:
        if os.path.isfile(filename):
            return filename
        else:
            return False
    except Exception as e:
        print(e)
        sys.exit()

def check_if_valid_ip(ip):
    try:
        ipaddress.ip_network(ip)
        return True
    except Exception as e:
        return False

organizationid = input('ORG_ID: ')
input_file = {
    "name": input('Input_file: '),
    'is_umbrella_list': False
}

while check_if_file_exists(input_file['name']) == False:
    print(f"Input_file {input_file} not found. Type in a valid file.")
    input_file['name'] = input('Input_file: ')

input_file['is_umbrella_list'] = check_if_umbrella_list(input_file['name'])

client_id = getpass.getpass('API_KEY: ')
client_secret = getpass.getpass('API_SECRET: ')

print('Do you wish to get a verbose report or a summary report? Verbose will include all DNS requests for each domain/url/ip during the last days up to a maximum of 5000 requests, while summary will get the amount of requests for each domain/url/ip.')

input_verbose = {
    'user_input': input('Verbose [y/n]: '),
    'mode': ''
}

while input_verbose['user_input'] not in ['y', 'n']:
    input_verbose['user_input'] = input('Verbose [y/n]: ')

if input_verbose['user_input'] == 'y':
    input_verbose['mode'] = 'verbose'
elif input_verbose['user_input'] == 'n':
    input_verbose['mode'] = 'summary'

class UmbrellaAPI:
    def __init__(self, url, ident, secret):
        self.url = url
        self.ident = ident
        self.secret = secret
        self.token = None

    def GetToken(self):
        auth = HTTPBasicAuth(self.ident, self.secret)
        client = BackendApplicationClient(client_id=self.ident)
        oauth = OAuth2Session(client=client)
        self.token = oauth.fetch_token(token_url=self.url, auth=auth)
        return self.token

try:
    api = UmbrellaAPI('https://management.api.umbrella.com/auth/v2/oauth2/token', client_id, client_secret)   
    token = api.GetToken()
except Exception as e:
    print('Unable to retrieve Umbrella token.')
    print(e)
    sys.exit()

worksheet_row = 1
workbook_name = f"Umbrella_activity_report_{input_verbose['mode']}_{datetime.now().strftime('%d%m%Y-%H%M%S')}.xlsx"
workbook = xlsxwriter.Workbook(workbook_name)
worksheet = workbook.add_worksheet()

if input_verbose['user_input'] == 'y':
    worksheet.write('A1', 'allapplications')
    worksheet.write('B1', 'allowedapplications')
    worksheet.write('C1', 'blockedapplications')
    worksheet.write('D1', 'categories')
    worksheet.write('E1', 'date')
    worksheet.write('F1', 'domain')
    worksheet.write('G1', 'externalip')
    worksheet.write('H1', 'identities')
    worksheet.write('I1', 'internalip')
    worksheet.write('J1', 'policycategories')
    worksheet.write('K1', 'querytype')
    worksheet.write('L1', 'returncode')
    worksheet.write('M1', 'threats')
    worksheet.write('N1', 'time')
    worksheet.write('O1', 'timestamp')
    worksheet.write('P1', 'type')
    worksheet.write('Q1', 'verdict')

    with open(input_file['name'], newline='', encoding='UTF-8-SIG') as csvfile:
        if input_file['is_umbrella_list']:
            reader = csv.DictReader(csvfile)
            for row in reader:
                query = san_query(row['destination'])
                if row.get('type') == 'domain':
                    query_type = 'domains'
                elif row.get('type') == 'url':
                    query_type = 'urls'
                elif row.get('type') == 'ip':
                    query_type = 'ip'
                else:
                    print(f"Unable to determine destination type for {row.get('destination', '(Unable to get destination.)')}")
                    continue
                
                print(f"Getting information for {query}")

                try:
                    api_headers = {'Authorization': f"Bearer {token['access_token']}"}
                    base_url = f"https://reports.api.umbrella.com/v2/"
                    url = f"{base_url}organizations/{organizationid}/activity?from=-30days&to=now&limit=5000&{query_type}={query}"
                    req = requests.get(url, headers=api_headers)
                    req_json = req.json()
                    activities = req_json.get('data', [])
                    if activities == []:
                        print(f"No activities found for {query}")
                        print("Responce from Umbrella: ", end="")
                        pprint(req_json)
                        continue
                    
                    for activity in activities:
                        worksheet.write(worksheet_row, 0, str(activity.get('allapplications', 'NA')))
                        worksheet.write(worksheet_row, 1, str(activity.get('allowedapplications', 'NA')))
                        worksheet.write(worksheet_row, 2, str(activity.get('blockedapplications', 'NA')))
                        worksheet.write(worksheet_row, 3, str(activity.get('categories', 'NA')))
                        worksheet.write(worksheet_row, 4, str(activity.get('date', 'NA')))
                        worksheet.write(worksheet_row, 5, str(activity.get('domain', 'NA')))
                        worksheet.write(worksheet_row, 6, str(activity.get('externalip', 'NA')))
                        worksheet.write(worksheet_row, 7, str(activity.get('identities', 'NA')))
                        worksheet.write(worksheet_row, 8, str(activity.get('internalip', 'NA')))
                        worksheet.write(worksheet_row, 9, str(activity.get('policycategories', 'NA')))
                        worksheet.write(worksheet_row, 10, str(activity.get('querytype', 'NA')))
                        worksheet.write(worksheet_row, 11, str(activity.get('returncode', 'NA')))
                        worksheet.write(worksheet_row, 12, str(activity.get('threats', 'NA')))
                        worksheet.write(worksheet_row, 13, str(activity.get('time', 'NA')))
                        worksheet.write(worksheet_row, 14, str(activity.get('timestamp', 'NA')))
                        worksheet.write(worksheet_row, 15, str(activity.get('type', 'NA')))
                        worksheet.write(worksheet_row, 16, str(activity.get('verdict', 'NA')))
                        worksheet_row += 1
                except Exception as e:
                    print(f"Unable to get information from Umbrella for {query} ")
                    continue

        else:
            reader = csv.reader(csvfile)
            for row in reader:
                query = san_query(row[0])
                if check_if_valid_ip(query):
                    query_type = 'ip'
                elif "/" in query or "&" in query:
                    query_type = 'urls'
                else:
                    query_type = 'domains'
                
                print(f"Getting information for {query}")

                try:
                    api_headers = {'Authorization': f"Bearer {token['access_token']}"}
                    base_url = f"https://reports.api.umbrella.com/v2/"
                    url = f"{base_url}organizations/{organizationid}/activity?from=-30days&to=now&limit=5000&{query_type}={query}"
                    req = requests.get(url, headers=api_headers)
                    req_json = req.json()
                    activities = req_json.get('data', [])
                    if activities == []:
                        print(f"No activities found for {query}")
                        print("Responce from Umbrella: ", end="")
                        pprint(req_json)
                        continue
                    
                    for activity in activities:
                        worksheet.write(worksheet_row, 0, str(activity.get('allapplications', 'NA')))
                        worksheet.write(worksheet_row, 1, str(activity.get('allowedapplications', 'NA')))
                        worksheet.write(worksheet_row, 2, str(activity.get('blockedapplications', 'NA')))
                        worksheet.write(worksheet_row, 3, str(activity.get('categories', 'NA')))
                        worksheet.write(worksheet_row, 4, str(activity.get('date', 'NA')))
                        worksheet.write(worksheet_row, 5, str(activity.get('domain', 'NA')))
                        worksheet.write(worksheet_row, 6, str(activity.get('externalip', 'NA')))
                        worksheet.write(worksheet_row, 7, str(activity.get('identities', 'NA')))
                        worksheet.write(worksheet_row, 8, str(activity.get('internalip', 'NA')))
                        worksheet.write(worksheet_row, 9, str(activity.get('policycategories', 'NA')))
                        worksheet.write(worksheet_row, 10, str(activity.get('querytype', 'NA')))
                        worksheet.write(worksheet_row, 11, str(activity.get('returncode', 'NA')))
                        worksheet.write(worksheet_row, 12, str(activity.get('threats', 'NA')))
                        worksheet.write(worksheet_row, 13, str(activity.get('time', 'NA')))
                        worksheet.write(worksheet_row, 14, str(activity.get('timestamp', 'NA')))
                        worksheet.write(worksheet_row, 15, str(activity.get('type', 'NA')))
                        worksheet.write(worksheet_row, 16, str(activity.get('verdict', 'NA')))
                        worksheet_row += 1
                except Exception as e:
                    print(f"Unable to get information from Umbrella for {query} ")
                    continue

if input_verbose['user_input'] == 'n':
    worksheet.write('A1', 'Destination')
    worksheet.write('B1', 'Requests')

    with open(input_file['name'], newline='', encoding='UTF-8-SIG') as csvfile:
        if input_file['is_umbrella_list']:
            reader = csv.DictReader(csvfile)
            for row in reader:
                query = san_query(row['destination'])
                worksheet.write(worksheet_row, 0, query)
                if row.get('type') == 'domain':
                    query_type = 'domains'
                elif row.get('type') == 'url':
                    query_type = 'urls'
                elif row.get('type') == 'ip':
                    query_type = 'ip'
                else:
                    print(f"Unable to determine destination type for {query}")
                    worksheet.write(worksheet_row, 1, f"Unable to determine destination type for {query}")
                    continue
                
                print(f"Getting information for {query}")

                try:
                    api_headers = {'Authorization': f"Bearer {token['access_token']}"}
                    base_url = f"https://reports.api.umbrella.com/v2/"
                    url = f"{base_url}organizations/{organizationid}/activity?from=-30days&to=now&limit=5000&{query_type}={query}"
                    req = requests.get(url, headers=api_headers)
                    req_json = req.json()
                    activities = req_json.get('data', [])
                    if activities == []:
                        print(f"No activities found for {query}")
                        print("Responce from Umbrella: ", end="")
                        pprint(req_json)
                    worksheet.write(worksheet_row, 1, int(len(activities)))

                except Exception as e:
                    worksheet.write(worksheet_row, 1, "Unable to get information from Umbrella for {query}")
                    print(f"Unable to get information from Umbrella for {query}")
                    
                worksheet_row += 1
        else:
            reader = csv.reader(csvfile)
            for row in reader:
                query = san_query(row[0])
                worksheet.write(worksheet_row, 0, query)
                if check_if_valid_ip(query):
                    query_type = 'ip'
                elif "/" in query or "&" in query:
                    query_type = 'urls'
                else:
                    query_type = 'domains'
                
                print(f"Getting information for {query}")

                try:
                    api_headers = {'Authorization': f"Bearer {token['access_token']}"}
                    base_url = f"https://reports.api.umbrella.com/v2/"
                    url = f"{base_url}organizations/{organizationid}/activity?from=-30days&to=now&limit=5000&{query_type}={query}"
                    req = requests.get(url, headers=api_headers)
                    req_json = req.json()
                    activities = req_json.get('data', [])
                    if activities == []:
                        print(f"No activities found for {query}")
                        print("Responce from Umbrella: ", end="")
                        pprint(req_json)
                    worksheet.write(worksheet_row, 1, int(len(activities)))

                except Exception as e:
                    worksheet.write(worksheet_row, 1, "Unable to get information from Umbrella for {query}")
                    print(f"Unable to get information from Umbrella for {query}")
                    
                worksheet_row += 1

workbook.close()
input(f"Spreadsheet created. See {workbook_name} for the result.")