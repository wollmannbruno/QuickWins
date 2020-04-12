#!/usr/bin/env python

from netmiko import ConnectHandler
from getpass import getpass
from datetime import datetime
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo


def main():

    print('\n')
    username = input('Username:')
    password = getpass()
    print('\n')

    primary = {
        'device_type': 'cisco_asa',
        'host': 'fwl-dc1-vpn-a',
        'username': username,
        'password': password,
    }

    secondary = {
        'device_type': 'cisco_asa',
        'host': 'fwl-dc0-inet-a',
        'username': username,
        'password': password,
    }

    now = datetime.now()
    tab_name = now.strftime('%Y|%m|%d_%H|%M|%S')
    results = []

    for firewall in [primary, secondary]:
        print('Gathering information from {}'.format(firewall['host']))
        results.append(firewall['host'])
        vpn_sessiondb = show_vpn_sessiondb(firewall)
        results.append(vpn_sessiondb)

    output_to_excel(tab_name, results)    


def show_vpn_sessiondb(device):
    '''
    '''

    net_connect = ConnectHandler(**device)
    output = net_connect.send_command('show vpn-sessiondb anyconnect', use_textfsm=True)
    net_connect.disconnect()
    return(output)


def output_to_excel(tab, data):
    '''
    '''

    PATH = r'S:\Cit\Operations\Network\AnyConnect'
    FILE = r'\AnyConnect.xlsx'
    excel_file = PATH + FILE

    print('Recording in spreadsheet {}'.format(excel_file))

    try:
        wb = load_workbook(filename = excel_file)
    except:
        wb = Workbook()
        ws = wb.active
        wb.remove(ws)

    ws = wb.create_sheet(tab, 0)

    ws.cell(1, 1, 'Firewall')
    column_number = {}
    # Using the first dictionary from the second item
    # in the list to generate the column headings
    headings = data[1][0].keys()
    number_of_columns = len(headings) + 1
    for element, heading in enumerate(headings, 2):
        # Populating the column header cells
        ws.cell(1, element, heading)
        # Creating a dictionary to store the column headers
        # with their column number because dictionaries are unordered
        column_number[heading] = element

    row_number = 1

    # Iterating through the list
    for item in data:
        if type(item) == str:
            hostname = item
        else:
            # Iterating through the list of dictionaries
            for row_data in item:
                row_number += 1
                ws.cell(row_number, 1, hostname)
                for column_heading in headings:
                    ws.cell(row_number, column_number.get(column_heading), row_data.get(column_heading))

    # table_ref = 'A1:Z{}'.format(row_number)
    # vpn_table = Table(displayName = 'VPN_Table', ref = table_ref)
    # ws.add_table(vpn_table)
    ws.freeze_panes = 'D2'
    wb.save(excel_file)
    print('There are {} rows and {} columns in the spreadsheet'.format(row_number, number_of_columns))


if __name__ == "__main__":
    main()
