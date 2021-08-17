from pymongo import MongoClient
from bson.objectid import ObjectId

client = MongoClient()
report = client.report

def insert_report(org_id, type, address, file_name, path):
    

    inspection = {
        'org_id': org_id,
        'property_id': None,
        'type': type,
        'address': address,
        'is_impoerted': False,
        'json_file_name': file_name,
        'path': path
    }
    inspections = report.inspections

    try:
        inspections.insert_one(inspection)
    except Exception as e:
        if e.code == 11000:
            print('Already exist in database')
        else:
            raise Exception(e)
    

def fetch_all_unimported_documents(org_id):
    return report.inspections.find({'property_id' : { '$ne': None }, 'is_impoerted': False, 'org_id': org_id})
    

def mark_as_imported(property_id):
    report.inspections.update_one({'property_id':property_id}, {'$set': {'is_impoerted': True}})

#temp
def fetch_all():
    return report.inspections.find({})
#temp    
def add_property_id(_id, property_id):
    report.inspections.update_one({'_id': ObjectId(_id)},{'$set':{'property_id':property_id}})
    
#client.close()
