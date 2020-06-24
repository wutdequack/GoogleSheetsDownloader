from __future__ import print_function
from collections import OrderedDict
from time import time
from datetime import datetime
import sys
import pickle
import os.path
import requests
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/drive.metadata.readonly']


def get_formatted_date():
    # Gets formatted Date
    return datetime.now().strftime("%d%b")


def enum_sn(range_of_sn, ignore_sn):
    """
    To create list of serial numbers
    :param range_of_sn: <list> list of serial numbers
    :param ignore_sn: <list> list of serial numbers to ignore
    :return: <list> final list of serial numbers
    """
    return sorted([sn for sn in range_of_sn if sn not in ignore_sn])


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
    dict_of_codes = OrderedDict()
    for url in list_of_urls:
        if not url.startswith("http://www.") or url.startswith("www.") or url.startswith("https://www."):
            url = "http://www." + url
        person_id = url.split("/")[-1]
        r = requests.get(url)
        dict_of_codes[person_id.upper()[:2]] = convert_unicode_to_str(r.url).split("/")[5]
    return dict_of_codes


def main(argv):
    """Shows basic usage of the Drive v3 API.
    Prints the names and ids of the first 10 files the user has access to.
    """
    print("[*] Welcome to Getting Google Sheets with wutdequack.")
    print("[*] Getting Credentials...")

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

    print("[*] Starting Google Drive Service API v3")
    service = build('drive', 'v3', credentials=creds)

    # Get Params
    if len(argv) != 4:
        print("Usage: getsheets.py <list of domainS seperated by commas> <serial num range> [serial num(s) to exclude, if any]")
        print("i.e. getsheets.py http://www.go.gov.sg/XXXX,http://www.go.gov.sg/XX 106-120 108,109")
        print("i.e. getsheets.py http://www.go.gov.sg/XXXX,http://www.go.gov.sg/XX 106-120 108")
        return None

    print("[*] Converting Domains into spreadsheetIds")

    # Enumerate list
    list_of_addresses = argv[1].split(",")
    dict_of_codes = get_codes(list_of_addresses)

    # Create Directory
    dir_name = "{}".format(time())
    os.mkdir(dir_name)
    accessToken = convert_unicode_to_str(creds.token)

    # Get list of S/Ns
    range_of_sn = range(int(argv[2].split("-")[0]), int(argv[2].split("-")[1]) + 1)
    range_of_sn = enum_sn(range_of_sn, map(int, argv[3].split(",")))

    if len(range_of_sn) != len(dict_of_codes):
        print("[!] Your Serial number range [{}] does not add up to the number of domains [{}] given...".format(len(range_of_sn), len(dict_of_codes)))
        print("[!] Check your params again!")
        return None

    print("[*] Enumerating through and forming Export URLs...")

    # Wah I am super tired. Will use a count as an index for now
    count = 0
    for person_id, spreadsheetId in dict_of_codes.items():
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
        file_loc = '{}/{}_{}_{}.pdf'.format(dir_name, get_formatted_date(), range_of_sn[count],person_id)
        with open(file_loc, 'wb') as saveFile:
            saveFile.write(r.content)
        print("[*] Created PDF File at {}".format(file_loc))
        count += 1

    print("[*] End of Program, Have a nice day!")


if __name__ == '__main__':
    main(sys.argv)
