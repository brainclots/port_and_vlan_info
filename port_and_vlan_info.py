# $language = "python"
# $interface = "1.0"

'''
Purpose:    Collect the interface status output on a switch and match it to
            MAC address info. Reference an external worksheet of OUIs
            and create formulae to display the manufacturer of each MAC.
Author:
            ___  ____ _ ____ _  _    _  _ _    ____ ___ ___
            |__] |__/ | |__| |\ |    |_/  |    |  |  |    /
            |__] |  \ | |  | | \|    | \_ |___ |__|  |   /__
            Brian.Klotz@nike.com

Version:    1.0
Date:       August 2017
'''

import os
from openpyxl import Workbook
import datetime

script_tab = crt.GetScriptTab()
script_tab.Screen.Synchronous = True
script_tab.Screen.IgnoreEscape = True


def main():
    if not crt.Session.Connected:
        crt.Dialog.MessageBox(
            "This script currently requires a valid connection to a "
            "Cisco distribution switch or other similar device.\n\n"
            "Please connect and then run this script again.")
        return

    screenRow = script_tab.Screen.CurrentRow
    screenCol = script_tab.Screen.CurrentColumn
    prompt = script_tab.Screen.Get(screenRow, 1, screenRow, screenCol).strip()
    if 'config' in prompt:
        crt.Dialog.MessageBox('Run script from user or priviliged exec only.')
        return
    switch_name = prompt[:-1]
    # switch_name = 'test'

    # Gather data
    script_tab.Screen.Send("term len 0\n")
    script_tab.Screen.WaitForString('\n')
    mac_output = CaptureOutputOfCommand('show mac add dyn | in Gi', prompt)
    sh_int_output = CaptureOutputOfCommand('show int status | in /', prompt)
    script_tab.Screen.Send("term no len\n")

    # Prep the output file
    filename = switch_name + '.xlsx'
    filename = os.path.join(os.environ['TMPDIR'], filename)
    wb = Workbook()
    wb.save(filename=filename)
    ws1 = wb.active

    # Prep MAC table output
    ws1.title = 'MAC_Table'
    ws1['A1'] = 'Port'
    ws1['B1'] = 'MAC Address'
    ws1['C1'] = 'Vlan'
    ws1['D1'] = 'Vendor'
    ws1.column_dimensions['A'].width = 13
    ws1.column_dimensions['B'].width = 13
    ws1.column_dimensions['C'].width = 13
    ws1.column_dimensions['D'].width = 35
    ws1.column_dimensions['E'].width = 18
    ws1['E1'] = datetime.datetime.now()
    mac_index = 1
    for row in mac_output.splitlines():
        if 'dynamic' in row.lower():
            mac_index += 1
            row = row.split()
            vlan = row[0]
            mac_addr = row[1]
            mac_interface = row[3]
            ws1['A' + str(mac_index)] = mac_interface
            ws1['B' + str(mac_index)] = mac_addr
            ws1['C' + str(mac_index)] = vlan
            ws1['D' + str(mac_index)] =\
                '=VLOOKUP(LEFT(B%s,7),\
            \'/Users/bklotz/Documents/OUI_Table.xlsx\'!Vendor_Table,2,FALSE)' \
            % mac_index

    # Prep show interface output
    ws2 = wb.create_sheet(title='Int_Status')
    ws2['A1'] = 'Port'
    ws2['B1'] = 'State'
    ws2['C1'] = 'VLAN'
    ws2['D1'] = 'MAC Address'
    ws2.column_dimensions['A'].width = 13
    ws2.column_dimensions['B'].width = 13
    ws2.column_dimensions['C'].width = 13
    ws2.column_dimensions['D'].width = 15
    ws2.column_dimensions['E'].width = 18
    ws2['E1'] = datetime.datetime.now()

    sh_int_index = 1  # Start index at 1 so that data starts in row 2
    for row in sh_int_output.splitlines():
        sh_int_index += 1
        row = row.split()
        port = row[0]
        if ('Po' not in port) and (len(row) == 6):  # 6 arguments means no description
            state = row[1]
            vlan = row[2]
        else:  # More arguments means a description is in place
            state = row[2]
            vlan = row[3]
        ws2['A' + str(sh_int_index)] = port
        ws2['B' + str(sh_int_index)] = state
        ws2['C' + str(sh_int_index)] = vlan
        ws2['D' + str(sh_int_index)] = (
            '=IF(ISNA(VLOOKUP(A{0},MAC_Table!$A:$D,2,False)),"",'
            'VLOOKUP(A{0},MAC_Table!$A:$D,2,False))'.format(sh_int_index)
        )
        ws2['E' + str(sh_int_index)] = (
            '=IF(ISNA(VLOOKUP(A{0},MAC_Table!$A:$D,4,FALSE)),"",'
            'VLOOKUP(A{0},MAC_Table!$A:$D,4,FALSE))'.format(sh_int_index)
        )
    ws2.auto_filter.ref = 'A1:D{0}'.format(sh_int_index)
    # Save spreadsheet
    wb.save(filename)
    # Open spreadsheet
    os.system('open %s' % filename)


def CaptureOutputOfCommand(command, prompt):
    output = ''
    script_tab.Screen.Send(command + '\n')
    script_tab.Screen.WaitForString('\n')
    output = script_tab.Screen.ReadString(prompt)
    return output


main()
