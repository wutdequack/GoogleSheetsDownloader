# Drive V3 Python Quickstart

Complete Step 1 described in the [Drive V3 Python Quickstart](
https://developers.google.com/drive/v3/web/quickstart/python)

Leave Most settings as Default.

Configure OAuth Client as Desktop app.

Download credentials.json and place it in same folder as the project.

Upon running code, your browser may prompt you for some stuff. Select the Google Account you did Step 1 with.

The results will be in a new directory starting with a lot of numbers (Epoch Time)

USE PYTHON 2.7!!!!!!!!
## Install

```
pip install -r requirements.txt
```

## Run

```
Usage: getsheets.py <list of domainS seperated by commas> <serial num range> [serial num(s) to exclude, if any]
Example 1: getsheets.py http://www.go.gov.sg/XXXX,http://www.go.gov.sg/XX 106-120 108,109
Example 2: getsheets.py http://www.go.gov.sg/XXXX,http://www.go.gov.sg/XX 106-120 108 
```

## Main Sources of Information

```
https://stackoverflow.com/questions/55731292/setting-pdf-export-options-in-google-drive-for-python
https://developers.google.com/drive/api/v3/quickstart/python
```

## Main Sources of Information

Last thing... I know its crappy. It is scrappy 30 mins of coding. Will sort out bugs and errors when there is the time.