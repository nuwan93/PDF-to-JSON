import sys
sys.path.insert(0, r'E:\files\work\project\assetowl\pdf_to_json\report_conversion')
from Report import Report

class Entry(Report):
    def __init__(self, type):
        Report.__init__(self, type)
        self.items = []

    #Attaching items to a room and attching that room to the room list
    def attach_room_and_items(self):
        self.room['items'] = self.items
        self.data['rooms'].append(self.room)
        self.items = []
        self.room = {}

    #check if room name is reapiting
    def is_same_room(self, romm_name):
        if 'title' in self.room:
            if self.room['title'] == romm_name:
                return True
            return False

    #If commnets are divided into multiple rows this will attach the comments
    def combine_comments_in_item(self,comment):
        previous_comment = self.items[len(self.items) - 1]['comment']
        self.items[len(self.items) - 1]['comment'] += f' {comment.value}'

    #creating a new item
    def add_item(self, title, clean, undamaged, working,  comment):
        condition  = {
            'isClean': self.check_conditions(clean.value),
            'isUndamaged': self.check_conditions(undamaged.value),
            'isWorking': self.check_conditions(working.value),
        }
        
        self.items.append({
            'title' : title.value,
            "type": "FIXTURE",
            'comment': str(comment.value) if comment.value != None else "",
            'condition': condition           
            })
