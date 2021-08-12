from pymongo import MongoClient

def insert_report(org_id, type, address, file_name, path):
    client = MongoClient()
    report = client.report

    inspection = {
        'org_id': org_id,
        'property_id': None,
        'type': type,
        'address': address,
        'is_impoerted': False,
        'file_name': file_name,
        'path': path
    }
    inspections = report.inspections
    inspections.insert_one(inspection)
    #client.close()
