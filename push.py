#!/usr/bin/env python3
"""Upload local files to Apps Script project and (optionally) deploy.

Usage:
    python3 push.py              # upload files only (instant via /dev URL)
    python3 push.py --deploy     # also create new version + update prod deployment
"""

import sys
from pathlib import Path

from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

SCRIPT_ID = '1WwzMNQSy5toYvm1LKaIXSCI4Gicf8zYcVXpUwnsfK22gLVTbsDqX7Ibu'
SCOPES = [
    'https://www.googleapis.com/auth/script.projects',
    'https://www.googleapis.com/auth/script.deployments',
]

ROOT = Path(__file__).parent
OAUTH_FILE = ROOT / 'oauth.json'
TOKEN_FILE = ROOT / 'token.json'

FILES = [
    ('Code.gs',         'Code',       'SERVER_JS'),
    ('Index.html',      'Index',      'HTML'),
    ('appsscript.json', 'appsscript', 'JSON'),
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


def upload_content(service):
    files = []
    for local_name, asset_name, asset_type in FILES:
        path = ROOT / local_name
        if not path.exists():
            print(f'   ! skip {local_name} (not found)')
            continue
        files.append({
            'name':   asset_name,
            'type':   asset_type,
            'source': path.read_text(encoding='utf-8'),
        })
        print(f'   + {local_name} -> {asset_name}.{asset_type.lower()}')

    service.projects().updateContent(scriptId=SCRIPT_ID, body={'files': files}).execute()
    print('OK files uploaded')


def deploy_new_version(service):
    version = service.projects().versions().create(
        scriptId=SCRIPT_ID,
        body={'description': 'auto-push'},
    ).execute()
    version_number = version['versionNumber']
    print(f'OK new version: {version_number}')

    deployments = service.projects().deployments().list(scriptId=SCRIPT_ID).execute()
    versioned = [
        d for d in deployments.get('deployments', [])
        if d.get('deploymentConfig', {}).get('versionNumber') is not None
    ]
    if not versioned:
        print('   ! no versioned deployment found - create one manually first')
        return

    target = versioned[0]
    config = target['deploymentConfig']
    new_config = {
        'versionNumber':    version_number,
        'manifestFileName': config.get('manifestFileName', 'appsscript'),
        'description':      config.get('description', ''),
    }
    service.projects().deployments().update(
        scriptId=SCRIPT_ID,
        deploymentId=target['deploymentId'],
        body={'deploymentConfig': new_config},
    ).execute()
    print(f'OK deployment {target["deploymentId"][:12]}... -> v{version_number}')


def main():
    if not OAUTH_FILE.exists():
        sys.exit(f'X oauth.json missing in {ROOT}')

    service = build('script', 'v1', credentials=credentials())
    upload_content(service)

    if '--deploy' in sys.argv:
        deploy_new_version(service)
    else:
        print('   (run with --deploy to update the public webapp URL)')


if __name__ == '__main__':
    main()
