import json
import glob
import sys
from datetime import datetime
import time
sys.path.insert(0, r'C:\Users\Nuwan\Downloads\pdf_to_json-20210617T052048Z-001\pdf_to_json\report_conversion')
from db import insert_report

class Report:

    files = None
    path = None
    org_Id = None

    def __init__(self, type):
        self.data = {
            'importMode': 'CREATE_ONLY',
	        'type': type,
            'inspectedAt': None
        }
        self.type = type
        self.data['rooms'] = []
        self.room = {}
        self.property_adress = ''

    def set_date(self,inspection_date, format):
        d = datetime.strptime(inspection_date, format).date()
        unix_time = time.mktime(d.timetuple())
        self.data['inspectedAt'] = unix_time

    @classmethod    
    def load_excel_files(self, path, org_id):
        Report.files = glob.glob(rf'{path}\*.xlsx')
        Report.path = path
        Report.org_Id = org_id

     #return json condition marker according to the report
    def check_conditions(self, condition_marker):
        if condition_marker == 'Y' or condition_marker == 'P':
            return 'YES'
        elif condition_marker == 'N' or condition_marker == 'O':
            return 'NO'
        else:
            return 'NA'

    def set_property_address(self, address):
        self.property_adress = address.replace("/","_")
    
    #Extract alt title from the title
    def get_alt_title(self, title):
        alt_title = None
        if '/' in title:
            title, alt_title = title.split('/')
            return title, alt_title
        elif '(' in title:
            title, alt_title = title.split('(')
            return title, alt_title[:-1]
        return title, alt_title

    #creating a new room
    def create_room(self, main_title, alt_title, comment):
         self.room['title'] = main_title
         self.room['altTitle'] = alt_title
         self.room['comment'] = comment
    
   
    #cretin a JSON file and writing 
    def create_json(self,file_name):
        with open(rf'{Report.path}\{file_name}.json', 'w', encoding='utf-8') as outfile:
            json.dump(self.data, outfile, ensure_ascii=False, indent=4)

    def insert_to_db(self, file_name):
        insert_report(Report.org_Id, self.type, self.property_adress, file_name, Report.path) 
