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
 - pip install openpyxl
 - pip install xlrd
 - pip install xlwt
 - pip install tablib

 - or if behind a proxy server: pip install --proxy="user:password@server:port" packagename
 - within Transpower: pip install --proxy="transpower\mulhollandd:password@tptpxy001.transpower.co.nz:8080" tablib    

TODO: 
 - so many things
 - sorting options on display and dump output?    
 - sort out guessing of Transpower standard design version 
"""

__author__ = "Daniel Mulholland"
__copyright__ = "Copyright 2015, Daniel Mulholland"
__credits__ = ["Kenneth Reitz https://github.com/kennethreitz/tablib"]
__license__ = "GPL"
__version__ = '0.01'
__maintainer__ = "Daniel Mulholland"
__hosted__ = "https://github.com/danyill/tp-setting-excel-tool/"
__email__ = "dan.mulholland@gmail.com"

# update this line if using from Python interactive
#__file__ = r'W:\Education\Current\pytooldev\tp-setting-excel-tool'

import sys
import os
import argparse
import glob
import re
import tablib
import openpyxl
import xlrd
import xlwt

BASE_PATH = os.path.dirname(os.path.realpath(__file__))
OUTPUT_FILE_NAME = "output"
OUTPUT_HEADERS = ['Filename','Setting Name','Val','Spreadsheet Reference']
EXCEL_FILE_REGEX = '(xls|xlsx|xlsm)$'


def main(arg=None):
    parser = argparse.ArgumentParser(
        description='Process individual or multiple RDB files and produce summary'\
            ' of results as a csv or xls file.',
        epilog='Enjoy. Bug reports and feature requests welcome. Feel free to build a GUI :-)',
        prefix_chars='-/')

    parser.add_argument('-o', choices=['csv','xlsx'],
                        help='Produce output as either comma separated values (csv) or as'\
                        ' a Micro$oft Excel .xls spreadsheet. If no output provided then'\
                        ' output is to the screen.')

    parser.add_argument('path', metavar='PATH|FILE', nargs=1, 
                       help='Go recursively go through path PATH. Redundant if FILE'\
                       ' with extension .rdb is used. When recursively called, only'\
                       ' searches for files with:' +  SR760_EXTENSION + '. Globbing is'\
                       ' allowed with the * and ? characters.')

    parser.add_argument('-s', '--screen', action="store_true",
                       help='Show output to screen')

    # Not implemented yet
    #parser.add_argument('-d', '--design', action="store_true",
    #                   help='Attempt to determine Transpower standard design version and' \
    #                   ' include this information in output')
                       
    parser.add_argument('settings', metavar='G:S', type=str, nargs='+',
                       help='Settings in the form of G:S where G is the group'\
                       ' and S is the SEL variable name. If G: is omitted the search' \
                       ' goes through all groups. Otherwise G should be the '\
                       ' group of interest. S should be the setting name ' \
                       ' e.g. OUT201.' \
                       ' Examples: G1:50P1P or G2:50P1P or 50P1P' \
                       ' '\
                       ' You can also get port settings using P:S'
                       ' Note: Applying a group for a non-grouped setting is unnecessary'\
                       ' and will prevent you from receiving results.')

    parser.add_argument('-v', '--version', action='version', version='%(prog)s ' + __version__)

    if arg == None:
        args = parser.parse_args()
    else:
        args = parser.parse_args(arg.split())
    
    # read in list of files
    files_to_do = return_file_paths(args.path, SR760_EXTENSION)
    
    # sort out the reference data for addresses to parameter matching
    lookup = {}
    with open(INPUT_DATA_FILE, mode='r') as csvfile:
        ref_d = csv.DictReader(csvfile)        
        for row in ref_d:
            key = row.pop('SR750/760 - V401') + ':' + row.pop('Info Name')
            # we will assume no duplicates and ensure input data correct
            #if key in result:
            # implement your duplicate row handling here
            #    pass
            lookup[key] = row        
    
    if files_to_do != []:
        process_760_files(files_to_do, args, lookup)
    else:
        print('Found nothing to do for path: ' + args.path[0])
        sys.exit()
        os.system("Pause")
    
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
    
def process_760_files(files_to_do, args, reference_data):
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

    if args.screen == True:
        display_info(parameter_info)

def extract_parameters(filename, s_parameters, reference_data):
    parameter_info=[]
    
    # read data
    with open(filename,'r') as f:
        data = f.read()

    # get input arguments in a sane format
    s_parameters = re.findall(INPUT_SPLIT_EXPRESSION,' '.join(s_parameters))

    # ^0x'17AA=([\w :+/\\()!,.\-_\\*]*)
    # [SETPOINT GROUP 1](.|\n)*^0x17AA=([\w :+/\\()!,.\-_\\*]*)
    result = None
    for sp in s_parameters:
        # there is only device information and setpoint data
        grouping = sp.split(':')
        val = re.search(r'^\[' \
            + grouping[0].upper() \
            + r'\]\r\n', data, flags=re.MULTILINE)
        #print sp + " " + str(val.end())
        if val is not None:  
            if sp.split(':')[0] == 'Device Information':
                result = re.search('^' + \
                            reference_data[sp]['ADDRESS'] + \
                             SR760_SETPOINT_EXPR_END, \
                             data[val.end():], flags=re.MULTILINE).group(1)
            else:
                pattern = SR760_SETPOINT_EXPR_START + \
                            reference_data[sp]['ADDRESS'] + \
                             SR760_SETPOINT_EXPR_END
                result = re.search(pattern, data[val.end():], 
                    flags=re.MULTILINE).group(1)
            #print [sp,result]
        else:
            print "Grouping term not found"
        
        if result <> None:
            filename = os.path.basename(filename)
            parameter_info.append([filename, sp, result])
            
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
        main(r'-o xlsx in \
        G1:TRIP G1:TR')
    else:
        main()
    os.system("Pause")
        
