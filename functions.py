from google.oauth2 import service_account  # grabs service account
from classes import SheetsInventoryApi

def introduction_screen():
    title = 'INVENTORY MANAGEMENT'
    print('-'* len(title))
    print('INVENTORY MANAGEMENT')
    print('-' * len(title))
    print('About: This software is meant to aid in the entry and formatting of inventory in '
          'the spreadsheet "Updated Lab Room Equipment Inventory" on the PLC'
          'Drive. THis software intends to speed up the process by autoformatting'
          ' the source spreadsheet and making entries faster and more efficient.')
    first_sent = 'INSTRUCTIONS: There are 4 types of entries to make. ' \
                 'The first will always be the Starting cabinet number' \
               ', followed by one of these types in the next line: \n'

    instruct = 'INSTRUCTIONS: There are 4 types of entries to make. The ' \
               'first will always be the Starting cabinet number' \
               ', followed by one of these types in the next line: \n' \
               '\n Type 1: "na" - to facilitate fast entries of empty cabinets, all you have to do' \
               'is type "na" and this empty entry will be made to the next cabinet number automatically.' \
               '\n Type 2: "equipment name, amount" - this is the standard format for an entry where you want to' \
               'stay within the same cabinet. Seperate each category by a comma' \
               '\n Type 3: "cabinet number, equipment name, amount" - ' \
               'Once you are ready to change cabinets, type the ' \
               'next entry like this instead' \
               '\n Type 4: "cabinet number, na" - this is not totally necessary since you can just type "na" and it ' \
               'will automaticallu append to next line, but it works the same nonetheless'
    print('_' * len(first_sent))
    print(instruct)
    print('-' * len(first_sent))
    print('KEY WORDS: Enter "quit" to quit the program or "new" to switch to a different room')

def remove_spaces(element):
    remove_spaces= element.split()
    return remove_spaces[0]

def check_entry(entry):
    """
       Type 1: na  or new or quit
       Type 2: equipment, amount --(default) needs no error checks
       Type 3: cab#, equip, amount  --(changes cab number)
       Type 4: cab#, na --(change cab number and na)

       """

    if len(entry) == 1:  # TYpe 1, new and quit handled in main
        ENTRY = entry[0].lower()
        if ENTRY != 'na':
            return False, 0
        else:
            return True, 1  # extra number is the type of entry

    elif len(entry) == 2:  # Type 2 and 4
        # check if it is type 4
        ele = remove_spaces(entry[1])

        if ele.lower() != 'na': # if first element is not cab#
            # then it is type 2
            return True, 2
        else:  # then na is 2nd option
            # type 4
            ele = remove_spaces(entry[0])
            return check_cabinet_entry(ele), 4  # if first value is valid cab# then return true

    elif len(entry) == 3:  # Type 3
        ele = remove_spaces(entry[0])
        return check_cabinet_entry(ele), 3

    else:  # length is not one of the above
        print('Input is not Valid! Try again.')
        return False, 0


def check_cabinet_entry(number_string):
    number_string = remove_spaces(number_string)
    try:
        int_check = int(number_string)
        return True
    except:
        #print('"{}" CANNOT be a Cabinet NUmber!'.format(number))
        return False


def prep_API():
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets',
              'https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    SERVICE_ACCOUNT_FILE = 'client_secret.json'
    credential = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

    return SCOPES, SERVICE_ACCOUNT_FILE, credential


def is_quitting(entry):
    entry = remove_spaces(entry[0]).lower()

    if entry == 'quit':
        return True

    return False


def is_changing_room(entry):
    if len(entry) == 1:
        entry = remove_spaces(entry[0]).lower()
        if entry.lower() == 'new':
            return True
    return False  # not changing rooms!


def set_cabinet_number(inventory):
    while True:
        cabinet_num = input('Enter Starting Cabinet Number: ')
        cabinet_num = remove_spaces(cabinet_num)
        if check_cabinet_entry(cabinet_num):
            inventory.set_cabinet_number(int(cabinet_num))
            break
        else:
            print('"{}" CANNOT be a Cabinet NUmber!'.format(cabinet_num))
            continue

def check_and_set_sheet(inventory, sheet_name):
    sheet_name = remove_spaces(sheet_name)
    while True:
        try:
            inventory.set_worksheet(sheet_name)
            break
        except:
            print('\nSheet Name does not exist!')
            sheet_name = input('Enter New Room Number: ')
            continue


def set_entry(inventory, sheet_name):
    """
          Type 1: na  or new or quit
          Type 2: equipment, amount --(default) needs no error checks
          Type 3: cab#, equip, amount  --(changes cab number)
          Type 4: cab#, na --(change cab number and na)

          """
    delimitter = ','
    first_entry = True
    i = 0
    while True:
        if i != 0:
            first_entry = False
            statement = "\nEntry {}: ".format(i+1)
        else:  # if i is zero
            first_entry = True
            statement = "\nEnter Cabinet# (ONLY if changing cabinets), " \
                        "Equipment name, and Amount each separated by a comma: "

        entry = input(statement).split(delimitter)
        entry_is_valid, entry_type = check_entry(entry)

        if len(entry) == 1:  # cases where entry is quit or new
            if is_quitting(entry):
                break  # breaks and ends program
            elif is_changing_room(entry):
                sheet_name = input('Enter New Room Number: ')
                check_and_set_sheet(inventory, sheet_name)
                set_cabinet_number(inventory)  # new starting cabinet number for room
                i = 0  # reset entries when changing rooms
                continue

        if entry_is_valid:  # if valid entry
            if entry_type == 3 or entry_type == 4:
                print('\nMoving to Cabinet {}...'.format(int(entry[0])))

            if entry_type == 1 and not first_entry: # if entry is 'na' then append new entry with na
                current_num = inventory.get_cabinet_number()
                inventory.set_cabinet_number(current_num + 1)

            check_and_set_sheet(inventory, sheet_name)
            inventory.store_entry(entry, entry_type)
            print('\nAdding Entry to cabinet {}...'.format(inventory.get_cabinet_number()))
            inventory.write_entry()
            i += 1
            continue

        # length invalid
        print('\nInput is not Valid! Try again.')
        continue

def delete_entry():

    pass

def replace_entry():

    pass


