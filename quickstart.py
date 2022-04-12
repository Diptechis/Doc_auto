from __future__ import print_function

import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from google.oauth2 import service_account

SCOPES = ['https://www.googleapis.com/auth/documents']
SERVICE_ACCOUNT_FILE = 'key.json'

creds = None
creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# If modifying these scopes, delete the file token.json.


# The ID of a sample document.
DOCUMENT_ID = '1UNZlibK2FZUxwpD3WnFa0i4NO-Wyg5J7ler4M7IWwDY'

service = build('docs', 'v1', credentials=creds)
document = service.documents().get(documentId=DOCUMENT_ID).execute()



