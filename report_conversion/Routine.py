from Report import Report

class Routine(Report):

    def __init__(self, created_date, type):
        Report.__init__(self, created_date, type)

    #Attaching room 
    def attach_room(self):
        self.data['rooms'].append(self.room)
        self.room = {}

    #If commnets are divided into multiple rows this will attach the comments
    def combine_comments_in_room(self,comment):
        self.room['comment'] += f' \n {comment.value}'