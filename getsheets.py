from __future__ import print_function
from time import time
import sys
import pickle
import os.path
import requests
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/drive.metadata.readonly']


def convert_unicode_to_str(unicode_str):
    """
    Convert unicode str to str str
    :param unicode_str: unicode str
    :return: str str
    """
    return unicode_str.encode("ascii", "ignore")


def get_codes(list_of_urls):
    """
    Given list of URLs get Google Sheet Document Codes
    :param list_of_urls: list of urls
    :return: dict of codes
    """
    dict_of_codes = {}
    for url in list_of_urls:
        if not url.startswith("http://www.") or url.startswith("www.") or url.startswith("https://www."):
            url = "http://www." + url
        person_id = url.split("/")[-1]
        r = requests.get(url)
        dict_of_codes[convert_unicode_to_str(r.url).split("/")[5]] = person_id
    return dict_of_codes


def main(argv):
    """Shows basic usage of the Drive v3 API.
    Prints the names and ids of the first 10 files the user has access to.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('drive', 'v3', credentials=creds)

    # Get Params
    if len(argv) != 2:
        print("Usage: getsheets.py <list of domains seperated by commas>")
        print("i.e. getsheets.py http://www.go.gov.sg/XXXX,http://www.go.gov.sg/XX")
        return None

    # Enumerate list
    list_of_addresses = argv[1].split(",")
    dict_of_codes = get_codes(list_of_addresses)

    # Create Directory
    dir_name = "{}".format(time())
    os.mkdir(dir_name)
    accessToken = convert_unicode_to_str(creds.token)

    for spreadsheetId, person_id in dict_of_codes.items():
        url = ('https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export?'
               + 'format=pdf'  # export as PDF
               + '&portrait=false'  # landscape
               + '&top_margin=0.00'  # Margin
               + '&bottom_margin=0.00'  # Margin
               + '&left_margin=0.00'  # Margin
               + '&right_margin=0.00'  # Margin
               + '&pagenum=RIGHT'  # Put page number to right of footer
               + '&access_token=' + accessToken)  # access token
        r = requests.get(url)
        with open('{}/{}.pdf'.format(dir_name, person_id), 'wb') as saveFile:
            saveFile.write(r.content)
        print('Done.')


if __name__ == '__main__':
    main(sys.argv)
