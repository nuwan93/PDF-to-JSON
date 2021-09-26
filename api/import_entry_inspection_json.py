#
# Copyright (C) 2021 AssetOwl Technologies Pty Ltd, all rights reserved.
#

import json

import requests
import sys
import logging

sys.path.insert(0, r'C:\Users\nuwan\Documents\Projects\AssetOwl\PDF-to-JSON\report_conversion')
import db

logging.basicConfig(filename='bulk_import.log' ,level=logging.DEBUG, format='%(asctime)s: %(levelname)s:%(message)s')

# base_url = 'https://app.pip.local'
#base_url = 'https://app.dev.internal.assetowl.com'
# base_url = 'https://app.test.internal.assetowl.com'
base_url = 'https://app.prod.assetowl.com'

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


access_token = login('un', "pass")['accessToken']

session.headers = {
    'Authorization': f'Bearer {access_token}'
}

org_id = 'c8f53258-7d9c-4f92-af49-28d36bfb541b'


documents = db.fetch_all_unimported_documents(org_id)

for data in documents:
    logging.info(f"*-*-*-*-*-*-*-*-*-*-*-* Started {data['address']} import *-*-*-*-*-*-*-*-*-*-*-*")
    try:
        with open(fr"{data['path']}\{data['json_file_name']}", 'r') as file:
            template_json = file.read()
    except Exception as e:
        logging.error(e)
        continue
    
    property_id = data['property_id']

    try:
        session.post(f'{base_url}/api/import/orgs/{org_id}/properties/{property_id}/inspections', data=template_json, headers={'Content-Type': 'application/json'}).raise_for_status()
    except Exception as e:
        logging.error(e)
        continue
    try:
        db.mark_as_imported(property_id)
        logging.info(f"Imported {data['address']} successfully")
    except Exception as e:
        logging.error(e)

