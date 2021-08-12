#
# Copyright (C) 2021 AssetOwl Technologies Pty Ltd, all rights reserved.
#

import json

import requests

# base_url = 'https://app.pip.local'
base_url = 'https://app.dev.internal.assetowl.com'
# base_url = 'https://app.test.internal.assetowl.com'
# base_url = 'https://app.prod.internal.assetowl.com'

session = requests.Session()
# session.verify = False  # when using local environment without valid SSL certificate


def login(username, password):
    login_json = {
        'username': username,
        'password': password,
        'deviceId': 'SCRIPT'
    }
    login_response = session.post(f'{base_url}/api/auth/login', data=json.dumps(login_json))
    login_response.raise_for_status()
    return login_response.json()


access_token = login('_TODO_USERNAME_', '_TODO_PASSWORD_')['accessToken']

session.headers = {
    'Authorization': f'Bearer {access_token}'
}


with open('example_import_entry_inspection.json', 'r') as file:
    template_json = file.read()

org_id = '_TODO_ORG_ID_'
property_id = '_TODO_PROPERTY_ID_'

session.post(f'{base_url}/api/import/orgs/{org_id}/properties/{property_id}/inspections', data=template_json, headers={'Content-Type': 'application/json'}).raise_for_status()

print('Import successful')
