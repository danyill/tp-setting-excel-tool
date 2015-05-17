#!/usr/bin/python

# NB: As per Debian Python policy #!/usr/bin/env python2 is not used here.

"""
tp-setting-excel-tool.py

A tool to (at least) extract information from Transpower setting 
spreadsheets.

Usage defined by running with option -h.

This tool can be run from the IDLE prompt using the main def.

Thoughtful ideas most welcome. 

Installation instructions (for Python *2.7.9*):
 - pip install xlrd
 - pip install tablib

 - or if behind a proxy server: pip install --proxy="user:password@server:port" packagename
 - within Transpower: pip install --proxy="transpower\mulhollandd:password@tptpxy001.transpower.co.nz:8080" tablib    
 
TODO: 
 - so many things
 - sorting options on display and dump output?    
 - sort out guessing of Transpower standard design version 
 - sort out dumping all parameters and argparse dependencies
 - sort out extraction of DNP data and port settings
 - pivoting settings
"""

__author__ = "Daniel Mulholland"
__copyright__ = "Copyright 2015, Daniel Mulholland"
__credits__ = ["Kenneth Reitz https://github.com/kennethreitz/tablib"]
__license__ = "GPL"
__version__ = '0.03'
__maintainer__ = "Daniel Mulholland"
__hosted__ = "https://github.com/danyill/tp-setting-excel-tool/"
__email__ = "dan.mulholland@gmail.com"

# update this line if using from Python interactive
#__file__ = r'W:\Education\Current\pytooldev\tp-setting-excel-tool'

import sys
import os
import string
import argparse
import glob
import regex
import tablib
import xlrd

# for password protected xls a different 
# strategy is required we need to have
# excel installed to decrypt the workbook
# import win32com.client

BASE_PATH = os.path.dirname(os.path.realpath(__file__))
OUTPUT_FILE_NAME = "output"
OUTPUT_HEADERS = ['Filename','Setting Name','Val','Spreadsheet Reference']
EXCEL_FILE_REGEX = '(xls|xlsx|xlsm)$'

PARAMETER_SEPARATOR = ':'

SEL_SEARCH_EXPR = {\
    'G1': [['SEL321 DISTANCE RELAY GROUP 1 SETTINGS',
            'RELAY GROUP 1 SETTINGS (SET 1)',
            'GROUP 1 SETTINGS'], \
           ['SEL321 DISTANCE RELAY GROUP 2 SETTINGS',
           'GROUP 2 IS SET IDENTICAL TO GROUP 1',
           'RELAY GROUP 2 SETTINGS (SET 2)',
           'GROUP 6 SETTINGS'] \
          ], \
    'G2': [['SEL321 DISTANCE RELAY GROUP 2 SETTINGS',
            'RELAY GROUP 2 SETTINGS (SET 2)'], \
           ['SEL321 DISTANCE RELAY GROUP 6 SETTINGS',
            'GROUP 6 IS SET IDENTICAL TO GROUP 1'
            'RELAY GROUP 6 SETTINGS (SET 6)'] \
          ], \
    'G6': [['SEL321 DISTANCE RELAY GROUP 6 SETTINGS',
            'RELAY GROUP 6 SETTINGS (SET 6)'], \
           ['Settings Valid Until:',
            'GROUP 6 SETTINGS'] \
          ], \
    }

# Structure of Transpower setting spreadsheets differs
# between the SEL-321, SEL-311C, SEL-351S and SEL-387
SETTINGS_PRINTOUT_SHEETS=['Main_Settings_Printout', 'Settings_Printout', 
    'Summary']
RELAY_TYPE_SHEETS = ['Common_Info_And_Settings', 'Global_Settings', 
   'Global_Setting','General Data']
REVISION_SHEETS = ['Revision Log', 'Revision', 'Revision_Log', 'Revisions']
OUTPUT_HEADERS = ['File','Setting Name','Val']
EXCEL_EXTENSION = 'XLS'

def main(arg=None):
    parser = argparse.ArgumentParser(
        description='Process individual or multiple Transpower SEL setting' \
            'spreadsheet files and produce summary of results as a csv or' \
            ' xls file.'\
            ' '\
            ' NOTE: Only processes .xls files. Not .xlsx or .xlsm!',
        epilog='Enjoy. Bug reports and feature requests welcome. Feel free to build a GUI :-)',
        prefix_chars='-/')

    parser.add_argument('-o', choices=['csv','xlsx'],
                        help='Produce output as either comma separated values (csv) or as'\
                        ' a Micro$oft Excel .xls spreadsheet. If no output provided then'\
                        ' output is to the screen.')

    parser.add_argument('path', metavar='PATH|FILE', nargs='+', 
                       help='Go recursively go through path PATH. Redundant if FILE'\
                       ' with extension ' + EXCEL_EXTENSION + ' is used. When '\
                       ' recursively called, only'\
                       ' searches for files with:' +  EXCEL_EXTENSION + '. Globbing is'\
                       ' allowed with the * and ? characters.')

    parser.add_argument('-c', '--console', action="store_true",
                       help='Show output to console')

    # Not implemented yet
    # parser.add_argument('-a', '--all', action="store_true",
    #                   help='Output all settings!')                       
                       
    # Not implemented yet
    # parser.add_argument('-d', '--design', action="store_true",
    #                   help='Attempt to determine Transpower standard design version and' \
    #                   ' include this information in output')
                       
    parser.add_argument('-s', '--settings', metavar='G:S', type=str, nargs='+',
                       help='Settings in the form of G:S where G is the group'\
                       ' and S is the SEL variable name. If G: is omitted the search' \
                       ' goes through all groups. Otherwise G should be the '\
                       ' group of interest. S should be the setting name ' \
                       ' e.g. OUT201.' \
                       ' Examples: G1:50P1P or G2:50P1P or 50P1P' \
                       ' '\
                       ' For the SEL-387 no grouping parameter is required'\
                       ' '\
                       ' Note: Applying a group for a non-grouped setting is unnecessary'\
                       ' and will prevent you from receiving results.'\
                       ' Special parameters are the following self-explanatory items:'\
                       ' REVISION')

    parser.add_argument('-v', '--version', action='version', version='%(prog)s ' + __version__)

    if arg == None:
        args = parser.parse_args()
    else:
        args = parser.parse_args(arg.split())
    
    # read in list of files
    files_to_do = return_file_paths([' '.join(args.path)], EXCEL_EXTENSION)

    # sort out the reference data for knowing where to search in the text string
    lookup = SEL_SEARCH_EXPR
    if files_to_do != []:
        process_xls_files(files_to_do, args, lookup)
    else:
        print('Found nothing to do for path: ' + args.path[0])
        sys.exit()
        raw_input("Press any key to exit...")
    
def return_file_paths(args_path, file_extension):
    paths_to_work_on = []
    for p in args_path:
        p = p.translate(None, ",\"")
        if not os.path.isabs(p):
            paths_to_work_on +=  glob.glob(os.path.join(BASE_PATH,p))
        else:
            paths_to_work_on += glob.glob(p)
    files_to_do = []
    # make a list of files to iterate over
    if paths_to_work_on != None:
        for p_or_f in paths_to_work_on:
            if os.path.isfile(p_or_f) == True:
                # add file to the list
                print os.path.normpath(p_or_f)
                files_to_do.append(os.path.normpath(p_or_f))
            elif os.path.isdir(p_or_f) == True:
                # walk about see what we can find
                files_to_do = walkabout(p_or_f, file_extension)
    return files_to_do        

def walkabout(p_or_f, file_extension):
    """ searches through a path p_or_f, picking up all files with EXTN
    returns these in an array.
    """
    return_files = []
    for root, dirs, files in os.walk(p_or_f, topdown=False):
        #print files
        for name in files:
            if (os.path.basename(name)[-3:]).upper() == file_extension:
                return_files.append(os.path.join(root,name))
    return return_files
    
def process_xls_files(files_to_do, args, reference_data):
    parameter_info = []
        
    for filename in files_to_do:      
        extracted_data = extract_parameters(filename, args.settings, reference_data)
        parameter_info += extracted_data

    # for exporting to Excel or CSV
    data = tablib.Dataset()    
    for k in parameter_info:
        data.append(k)
    data.headers = OUTPUT_HEADERS

    # don't overwrite existing file
    name = OUTPUT_FILE_NAME 
    if args.o == 'csv' or args.o == 'xlsx': 
        # this is stupid and klunky
        while os.path.exists(name + '.csv') or os.path.exists(name + '.xlsx'):
            name += '_'        

    # write data
    if args.o == None:
        pass
    elif args.o == 'csv':
        with open(name + '.csv','wb') as output:
            output.write(data.csv)
    elif args.o == 'xlsx':
        with open(name + '.xlsx','wb') as output:
            output.write(data.xlsx)

    if args.console == True:
        display_info(parameter_info)

def in_both_lists(a, b):
    return [i for i in a if i in b]

def get_relay_type(workbook):
    worksheets = workbook.sheet_names()
    rtypes = in_both_lists(worksheets, RELAY_TYPE_SHEETS)
    return workbook.sheet_by_name(rtypes[0]).cell(0,0).value.split(' ')[0]
    
def find_between_rows(parameter, worksheet):
    grouper = None
    sp = None
    in_band = None
    
    if parameter.find(PARAMETER_SEPARATOR) != -1:
        grouper = parameter.split(PARAMETER_SEPARATOR)[0]
        sp = parameter.split(PARAMETER_SEPARATOR)[1]
        in_band = False
    else:
        sp = parameter
        in_band = True

    num_rows = worksheet.nrows - 1
    curr_row = -1
    while curr_row < num_rows:
        curr_row += 1
        row = worksheet.row(curr_row)
        for idx, r in enumerate(row):
            if grouper != None and r.value in SEL_SEARCH_EXPR[grouper][0]:
                in_band = True
            elif grouper != None and r.value in SEL_SEARCH_EXPR[grouper][1]:
                in_band = False
            
            if (sp + '=' == r.value) and in_band == True:
                return row[idx+1].value                
            elif (sp == r.value and row[idx+1].value == '=' and
                in_band == True):
                return row[idx+2].value
    return None

def extract_parameters(filename, settings, reference_data):
    parameter_info=[]

    # read data
    try:
        workbook = xlrd.open_workbook(filename)
    except xlrd.XLRDError:
        print 'Could not read spreadsheet: ' + filename + '. Possibly encrypted?'
        parameter_info.append([filename, 'FAIL', 'Could not read spreadsheet. Possibly encrypted.'])
        return parameter_info
    except:
        print 'Could not read spreadsheet: ' + filename + '. Possibly damaged?'
        parameter_info.append([filename, 'FAIL', 'Could not read spreadsheet. Possibly damaged.'])
        return parameter_info
        
    try:
        #need to add code for handling encrypted workbooks
        #either need to port e.g. the libreoffice code or do a win32com application 
        #excel = win32com.client.Dispatch('Excel.Application')
        #workbook = excel.Workbooks.open(r'c:\mybook.xls', 'password')
        #workbook.SaveAs('unencrypted.xls')
        #http://stackoverflow.com/questions/22789951/xlrd-error-workbook-is-encrypted-python-3-2-3

        # check to see that we have the following:
        # - a sheet with a revision log
        # - a main sheet from which we'll extract the relay type
        # - a settings printout/summary sheet from which we'll extract settings
        worksheets = workbook.sheet_names()
        revision_sheet = in_both_lists(worksheets, REVISION_SHEETS)
        setting_sheet = in_both_lists(worksheets, SETTINGS_PRINTOUT_SHEETS)
        rtype_sheet = in_both_lists(worksheets, RELAY_TYPE_SHEETS)
        
        fn = os.path.basename(filename)
        
        if revision_sheet and setting_sheet and rtype_sheet:
            parameter_info.append([fn, 'Relay', get_relay_type(workbook)])

            settings_sheet = workbook.sheet_by_name(setting_sheet[0])
            for parameter in settings:            
                result = find_between_rows(parameter,settings_sheet)
                if result <> None:
                    parameter_info.append([fn, parameter, result])
                else:
                    parameter_info.append([fn, parameter, 'Not Found'])
        else:
            print 'Not a valid setting spreadsheet:' + filename
        
    except:
        print 'Well that was a fail'
        parameter_info.append([filename, 'FAIL', 'Could not read spreadsheet.'])
    return parameter_info
    
def display_info(parameter_info):
    lengths = []
    # first pass to determine column widths:
    for line in parameter_info:
        for index,element in enumerate(line):
            try:
                lengths[index] = max(lengths[index], len(element))
            except IndexError:
                lengths.append(len(element))
    
    parameter_info.insert(0,OUTPUT_HEADERS)
    # now display in columns            
    for line in parameter_info:
        display_line = '' 
        for index,element in enumerate(line):
            display_line += element.ljust(lengths[index]+2,' ')
        print display_line

if __name__ == '__main__': 
    if len(sys.argv) == 1 :
        main(r'-o xlsx "/home/mulhollandd/Downloads/playing/" --settings TID RID G1:TR G1:81D1P G1:81D1D G1:81D2P G1:81D2P G1:E81')           
    else:
        main()
    raw_input("Press any key to exit...")
        

