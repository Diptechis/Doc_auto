import os
import datetime
import time
import io
from googleapiclient.http import MediaIoBaseUpload
from Google import Create_Service

CLIENT_SECRET_FILE = r'client_secrets.json'
API_SERVICE_NAME = 'docs'
API_VERSION = 'v1'
SCOPES = ['https://www.googleapis.com/auth/documents',
          'https://www.googleapis.com/auth/drive']

"""
Step 1. Create Google API Service Instances
"""
# Google Docs instance
service_docs = Create_Service(
    CLIENT_SECRET_FILE,
    'docs', 'v1',
    ['https://www.googleapis.com/auth/documents',
     'https://www.googleapis.com/auth/drive']
)
time.sleep(2)

# Google Drive instance
service_drive = Create_Service(
    CLIENT_SECRET_FILE,
    'drive',
    'v3',
    ['https://www.googleapis.com/auth/drive']
)
time.sleep(2)

# Google Sheets instance
service_sheets = Create_Service(
    CLIENT_SECRET_FILE,
    'sheets',
    'v4',
    ['https://www.googleapis.com/auth/spreadsheets']
)
time.sleep(2)

temp_doc_id = '1wwNIQmhmR0RCZN9gmAFByieHqtmvW-qSERGXnmUTSVA'
google_sh_id = '1anf8-d1iRCzrS-yGUgHP9kVgEN7lL4aAYJlGRUcAKFs'
google_drive_id = '1w5N1_X-9sSjjlOw_4K2VcY7TyElKYmJY'

responses = {}

"""
Step 2: Load record
"""

worksheet_name = "Resume"
responses['sheets'] = service_sheets.spreadsheets().values().get(
    spreadsheetId=google_sh_id,
    range=worksheet_name,
    majorDimension='ROWS',
).execute()

columns = responses['sheets']['values'][0]
records = responses['sheets']['values'][1:]

"""
Step 3: Iterate Each Record and Perform Mail Merge
"""


def mapping(merge_field, value=''):
    json_representation = {
        'replaceAllText': {
            'replaceText': value,
            'containsText': {
                'matchCase': 'True',
                'text': "{{{{{0}}}}}".format(merge_field)
            }
        }
    }
    return json_representation


for record in records:
    print("Processing record {0}".format(record[0]))

    # copy template doc file
    docu_file = 'Resume for {0}'.format(record[0])

    responses['docs'] = service_drive.files().copy(
        fileId=temp_doc_id,
        body= {
            'parents': [google_drive_id],
            'name': docu_file
        }

    ).execute()
    docu_id = responses['docs']['id']

    # Update Google Docs
    merge_field_info = [mapping(columns[ind], value) for ind, value in enumerate(record)]

    service_docs.documents().batchUpdate(
        documentId=docu_id,
        body={
            'requests': merge_field_info
        }
    ).execute()

    # Exporting in pdf
    #
    PDF_MIME_TYPE = 'application/pdf'
    byteString = service_drive.files().export(
        fileId=docu_id,
        mimeType=PDF_MIME_TYPE
    ).execute()

    media_object = MediaIoBaseUpload(io.BytesIO(byteString), mimetype=PDF_MIME_TYPE)
    service_drive.files().create(
        media_body=media_object,
        body={
            'parents': [google_drive_id],
            'name': '{0} (PDF).pdf'.format(docu_file)
        }
    ).execute()
print("Task Completed")
