from google.oauth2 import service_account  # grabs service account
from classes import SheetsInventoryApi
from functions import *


def main():
    introduction_screen()

    SCOPES, SERVICE_ACCOUNT_FILE, credential = prep_API()

    inventory = SheetsInventoryApi(SCOPES, SERVICE_ACCOUNT_FILE, credential)

    sheet_name = input('\nEnter Room Number: ')
    check_and_set_sheet(inventory, sheet_name)

    set_cabinet_number(inventory)  # preliminary

    set_entry(inventory, sheet_name)


main()
input('\nPress any key to continue...')