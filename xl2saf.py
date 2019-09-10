#!/usr/bin/python
'''
Name: xl2saf.py
Author: Saiful Amin <saiful@semanticconsulting.com>
Description: Convert Excel spreadsheet into DSpace Simple Archive Format (SAF)

'''

import os, sys, shutil, re
import lxml.etree as et
from optparse import OptionParser
from openpyxl import load_workbook
from utils import parse_header, get_list
from datetime import datetime

# Process args
parser = OptionParser()
parser.add_option("-f", "--input_file", dest="xlsx_file", help="Input XLSX file", metavar="FILE")
parser.add_option("-b", "--base_dir", dest="base_dir", default=None, help="Base directory for full-text files")
parser.add_option("-d", "--items_dir", dest="items_dir", default='import', help="Output directory for converted files")
(options, args) = parser.parse_args()

# Check the mandatory option
if options.xlsx_file is None:
    parser.print_help()
    sys.exit()

# Prepare Items Dir
if not os.path.exists(options.items_dir):
    os.mkdir(options.items_dir)

# Global dict for parsed header row
md_header = {}

finished = 0
# Open the Spreadsheet file
wb = load_workbook(filename=options.xlsx_file, read_only=True)
for sheet in wb:

    # Check if the header row is acceptable
    if re.match(r'dc\.', str(sheet['A1'].value)) or re.match('filenames', str(sheet['A1'].value)):
        print('\n' + str(sheet.title) + ': found ' + str(sheet.max_row) + ' row(s).')
    else:
        print('\n' + str(sheet.title) + ': Skipped, does not contain valid header row.')
        continue

    # Looping the rows
    for row_key, row in enumerate(sheet.rows):

        # Parse the header row into global dict
        if row_key == 0:
            md_header = parse_header(row)
            continue
        
        # Folder & file name for each row
        item_dir = os.path.join(options.items_dir,'item_' + "%03d" % (row_key))
        if not os.path.exists(item_dir):
            os.mkdir(item_dir)
        xml_file = os.path.join(item_dir, 'dublin_core.xml')

        # Root element
        data = et.Element('dublin_core', attrib={"schema": "dc"})

        fulltext_files = list()

        # Go through the columns data
        for col_key, cell_data in enumerate(row):
            element     = md_header[col_key]['element']
            qualifier   = md_header[col_key]['qualifier']
            delimit     = md_header[col_key]['delimit']
            content     = cell_data.value

            if not element or not content:
                continue
            
            if element == 'filenames':
                if delimit:
                    fulltext_files = get_list(content, delimit)
                else:
                    fulltext_files.append(content)
                continue
            
            if element == 'date':
                content = datetime.strftime(content, '%Y-%m-%d')
            
            # Handle multiple values
            if delimit:
                multi_value = get_list(content, delimit)
                for value in multi_value:
                    dcvalue = et.Element('dcvalue', attrib={"element": element, "qualifier": qualifier})
                    dcvalue.text = value
                    data.append(dcvalue)
            else:
                dcvalue = et.Element('dcvalue', attrib={"element": element, "qualifier": qualifier})
                dcvalue.text = content
                data.append(dcvalue)

        # create a new XML file with the results
        xmldata = et.tostring(data, encoding='UTF-8', method='xml', xml_declaration=True, pretty_print=True)
        item_dc = open(xml_file, "wb")
        item_dc.write(xmldata)
        finished += 1

        # Copy fulltext files and update content manifest
        if fulltext_files:
            if options.base_dir is None:
                print('--base_dir for fulltext files missing.')
                parser.print_help()
                sys.exit()
            contents_file = open(os.path.join(item_dir, 'contents'), "w")
            for f in fulltext_files:
                fpath = os.path.join(options.base_dir, f)
                if os.path.isfile(fpath):
                    shutil.copy(fpath, item_dir)
                    contents_file.write(f + "\tbundle:ORIGINAL\n")
                else:
                    print('Row '+ str(row_key) + ':\tMissing file: ' + str(fpath))

print('\nPrepared ' + str(finished) + ' records in \'' + options.items_dir + '\' folder.' )
