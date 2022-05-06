import gspread
import pandas as pd
from google.oauth2 import service_account  # grabs service account

class SheetsInventoryApi:

    # Class Variables
    # API Preliminaries
    client = None
    sheet = None
    worksheet = None
    data = None
    df = None
    # Entries
    room_name = None
    entry_type = 0

    cabinet_number = None
    equipment_name = None
    quantity = None
    # Formatting data
    max = 0
    sheet_id = None
    position = 0
    pandas_position = 0
    merge_count = 0
    appending = None

    def __init__(self, scopes, service_acc, credentials):
        self.scopes = scopes
        self.service_acc = service_acc
        self.credentials = credentials
        self.client = gspread.authorize(credentials)
        self.sheet = self.client.open('Updated Lab Room Equipment Inventory')

    # PUBLIC METHODS
    def set_worksheet(self, worksheet_name):
        self.worksheet = self.sheet.worksheet(worksheet_name)
        self.__set_sheet_data()
        self.__set_dataframe()
        self.room_name = worksheet_name


    def set_cabinet_number(self, number):
        self.cabinet_number = number

    def get_cabinet_number(self):
        return self.cabinet_number

    # Entry analysis (Error checks already done)
    def store_entry(self, arr, TYPE):
        # only correct values of following type are allowed here
        """
           Type 1: na
           Type 2: equipment, amount (default)
           Type 3: cab#, equip, amount  (changes cab number)
           Type 4: cab# na (change cab number and na)

           """
        self.entry_type = TYPE
        if TYPE == 1:  # next cabinet gets na also
            self.equipment_name = arr[0].upper()  # cabinet number already filled in
        elif TYPE == 2:
            # cabinet number already filled in
            self.equipment_name = arr[0]
            self.quantity = arr[1]
        elif TYPE == 3:  # cabinet number is already set for every other type!
            self.cabinet_number = int(arr[0])  # new cabinet number!
            self.equipment_name = arr[1]
            self.quantity = arr[2]
        elif TYPE == 4:
            self.cabinet_number = int(arr[0])
            self.equipment_name = arr[1].upper()

    def write_entry(self):
        ent = None
        # establishes max cab number and position of cabinet number on sheets and on  dataframe
        self.__set_start_end_pos()

        if (self.entry_type == 1) or (self.entry_type == 4):
            ent = [[self.cabinet_number, self.equipment_name, '']]
        elif (self.entry_type == 2) or (self.entry_type == 3):
            ent = [[self.cabinet_number, self.equipment_name, self.quantity]]

        # case for empty row, just want to update it
        if (self.merge_count) == 1 and (self.df['Equipment'][self.pandas_position] == ''):
            self.worksheet.update('A{}:C{}'.format(self.position, self.position), ent)

        # case for non empty row, want to insert row
        elif not self.appending and (not self.df.empty):
            self.worksheet.insert_row(ent[0], self.position)  # will always add row above first row of cabinet number
            self.__merge()

        # case for row at the end.
        else:  # Will fill in previous values  until we get to the value you want!
            diff = self.cabinet_number - self.max
            if diff != 1:
                i = 1
                while diff != 1:
                    self.worksheet.append_row([self.max + i, '', ''])  # empty row
                    i += 1
                    diff -= 1
            self.worksheet.append_row(ent[0])

    def print_data(self):
        print(self.df)

    ###############################################################
    # PRIVATE METHODS

    def __set_sheet_id(self):
        self.sheet_id = self.sheet.worksheet(self.room_name)._properties['sheetId']

    def __set_sheet_data(self):
        self.data = self.worksheet.get_all_records()

    def __set_dataframe(self):
        self.df = pd.DataFrame.from_dict(self.data)

    def __max_cabinet_number(self):
        self.max = 0
        empty_sheet = False
        #print(len(self.df['Cabinet']))
        if self.df.empty:
            empty_sheet = True
        # now max will be zero
        if not empty_sheet:
            for num in self.df['Cabinet']:
                if type(num) == int:
                    if self.max < num:
                        self.max = num

    def __set_start_end_pos(self): # private method
        # Grabbing start and end positions from the sheet cabinet number and calculating amount of rows to merge
        self.position = 2  # + 2 to account for index start at 0 and not including titles for the index
        self.appending = False
        self.pandas_position = 0
        self.__max_cabinet_number()
        if not self.df.empty:
            for i, num in enumerate(self.df['Cabinet']):
                if self.cabinet_number > self.max:  # if this is the case we want to break immidiately
                    self.appending = True  # so we can skip later code
                    break

                if str(num) == str(self.cabinet_number):
                    # print('Matched ' + num + ' with' + str(cabinet_number))
                    self.position += i  # adds index of cab number in pandas to poistion to give postion in sheets
                    self.pandas_position = i
                    break

        pandas_index = self.pandas_position
        next_cab = self.cabinet_number + 1
        self.merge_count = -1

        if not self.appending and (not self.df.empty):
            while True:
                self.merge_count += 1
                # adding new row at the end:
                if pandas_index == len(self.df['Cabinet']):
                    self.merge_count = 1
                    break
                if str(self.df['Cabinet'][pandas_index]) == str(next_cab):
                    break
                pandas_index += 1

    def __merge(self):
        self.__set_sheet_id()
        while True:  # try and catch is for when merge count is insifficient and i need to aggregate
            try:    # prevents program from crashing if merging error and instead fixes it by incrementing merge_count
                body = {
                        "requests": [
                            {
                                "mergeCells": {
                                    "mergeType": "MERGE_COLUMNS",
                                    "range": {  # In this sample script, all cells of "A1:C3" of "Sheet1" are merged.
                                        "sheetId": self.sheet_id,
                                        "startRowIndex": self.position - 1,
                                        "endRowIndex": self.position + self.merge_count,
                                        "startColumnIndex": 0,
                                        "endColumnIndex": 1
                                    }
                                }
                            }
                        ]
                    }
                #print('merging from sheet row, ', self.position-1, ' to ', self.position + self.merge_count)
                self.sheet.batch_update(body)
                break
            except:
                self.merge_count += 1
                continue









