#!/usr/bin/env python3
"""Run createInstallableTrigger() in the GAS project via the Apps Script API."""

import json, sys
from pathlib import Path

from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import google.auth.transport.requests
import urllib.request

SCRIPT_ID  = '1WwzMNQSy5toYvm1LKaIXSCI4Gicf8zYcVXpUwnsfK22gLVTbsDqX7Ibu'
ROOT       = Path(__file__).parent
OAUTH_FILE = ROOT / 'oauth.json'
TOKEN_FILE = ROOT / 'trigger_token.json'   # separate cache to avoid interfering with push.py

SCOPES = [
    'https://www.googleapis.com/auth/script.projects',
    'https://www.googleapis.com/auth/script.deployments',
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/script.scriptapp',
    'https://www.googleapis.com/auth/calendar',
    'https://www.googleapis.com/auth/documents',
    'https://www.googleapis.com/auth/drive',
]

def credentials():
    creds = None
    if TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(OAUTH_FILE), SCOPES)
            creds = flow.run_local_server(port=0)
        TOKEN_FILE.write_text(creds.to_json())
    return creds

def run_function(creds, function_name):
    url = f'https://script.googleapis.com/v1/scripts/{SCRIPT_ID}:run'
    body = json.dumps({'function': function_name, 'devMode': True}).encode()
    req = urllib.request.Request(url, data=body, method='POST')
    req.add_header('Authorization', 'Bearer ' + creds.token)
    req.add_header('Content-Type', 'application/json')
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            data = json.loads(r.read())
            if 'error' in data:
                print(f'   ! {function_name} errore: {data["error"]}')
            else:
                print(f'   OK {function_name} completato')
            return data
    except urllib.error.HTTPError as e:
        body = e.read().decode()
        print(f'   ! HTTP {e.code}: {body[:300]}')
        return None

def main():
    if not OAUTH_FILE.exists():
        sys.exit(f'X oauth.json mancante in {ROOT}')
    print('Autenticazione...')
    creds = credentials()
    print(f'Token: {creds.token[:20]}...')
    print('Esecuzione createInstallableTrigger()...')
    result = run_function(creds, 'createInstallableTrigger')
    if result is None:
        print('\nSe vedi un errore "script has not been deployed", devi:')
        print('1. Aprire il progetto Apps Script')
        print('2. Deploy > Nuova distribuzione > Tipo: Eseguibile API')
        print('3. Ripetere questo script')

if __name__ == '__main__':
    main()
