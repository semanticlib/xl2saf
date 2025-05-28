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
        # row_key is 0 for header, 1 for first data row, etc.
        # Excel row number is row_key + 1. Item folder uses row_key for 0xx numbering after header.

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
            if col_key not in md_header:
                print(f"Warning: Item item_{row_key:03d} (Excel Row {row_key + 1}): Missing header for column index {col_key}. Skipping cell.")
                continue

            element_info = md_header[col_key]
            element     = element_info['element']
            qualifier   = element_info['qualifier']
            delimit     = element_info['delimit']
            
            raw_content = cell_data.value

            if not element: # Skip if element name from header is empty
                continue
            
            # Handle special columns for files first
            if element == 'filenames':
                if raw_content is not None:
                    content_str = str(raw_content)
                    files_from_cell = get_list(content_str, delimit) # delimit can be ""
                    fulltext_files.extend(files_from_cell)
                continue # Move to the next cell/column

            elif element == 'foldername':
                if raw_content is not None:
                    content_str = str(raw_content)
                    folder_names_in_cell = get_list(content_str, delimit) # delimit can be ""

                    for folder_name in folder_names_in_cell: # folder_name is already stripped and non-empty by get_list
                        if not folder_name: # Should be caught by get_list, but as a safeguard
                            continue

                        if options.base_dir is None:
                            print(f"Warning: Item item_{row_key:03d} (Excel Row {row_key + 1}): --base_dir not specified. Cannot process folder '{folder_name}' from 'foldername' column.")
                            continue 

                        actual_folder_path = os.path.join(options.base_dir, folder_name)
                        if os.path.isdir(actual_folder_path):
                            for item_in_folder in os.listdir(actual_folder_path):
                                item_full_path_in_source = os.path.join(actual_folder_path, item_in_folder)
                                if os.path.isfile(item_full_path_in_source):
                                    # Path relative to base_dir for consistent handling in copy stage
                                    relative_file_path_to_basedir = os.path.join(folder_name, item_in_folder)
                                    fulltext_files.append(relative_file_path_to_basedir)
                        else:
                            print(f"Warning: Item item_{row_key:03d} (Excel Row {row_key + 1}): Folder '{folder_name}' (path: '{actual_folder_path}') in 'foldername' column not found or is not a directory.")
                continue # Move to the next cell/column
            
            # Process as metadata field if not a special column and content exists
            if raw_content is None:
                continue

            content_for_meta = raw_content

            if element == 'date':
                if isinstance(content_for_meta, datetime):
                    content_for_meta = datetime.strftime(content_for_meta, '%Y-%m-%d')
                else:
                    # If it's not a datetime object, use its string representation.
                    print(f"Warning: Item item_{row_key:03d} (Excel Row {row_key + 1}), Column '{md_header[col_key]['element']}.{md_header[col_key]['qualifier']}': Content '{content_for_meta}' is not a datetime object. Using its string value.")
                    content_for_meta = str(content_for_meta)
            
            content_str_meta = str(content_for_meta)

            # Handle multiple values for metadata using the improved get_list
            values_for_meta = get_list(content_str_meta, delimit) # delimit can be ""
            for value in values_for_meta: # value is already stripped and non-empty
                if value: # Ensure value is not empty
                    dcvalue = et.Element('dcvalue', attrib={"element": element, "qualifier": qualifier})
                    dcvalue.text = value
                    data.append(dcvalue)

        # create a new XML file with the results
        xmldata = et.tostring(data, encoding='UTF-8', method='xml', xml_declaration=True, pretty_print=True)
        with open(xml_file, "wb") as item_dc:
            item_dc.write(xmldata)
        finished += 1

        # Copy fulltext files and update content manifest
        if fulltext_files:
            if options.base_dir is None:
                print(f"Error: Item item_{row_key:03d} (Excel Row {row_key + 1}): --base_dir for fulltext files missing. This is required when 'filenames' or 'foldername' columns result in files to be processed.")
                parser.print_help()
                sys.exit()
            with open(os.path.join(item_dir, 'contents'), "w") as contents_file:
                for f_raw in fulltext_files:
                    f = str(f_raw).strip() # Ensure string and strip
                    if not f: continue

                    fpath = os.path.join(options.base_dir, f)
                    if os.path.isfile(fpath):
                        shutil.copy(fpath, item_dir)
                        # The 'contents' file lists filenames as they appear in the item directory (i.e., basenames)
                        contents_file.write(os.path.basename(f) + "\tbundle:ORIGINAL\n")
                    else:
                        print(f"Warning: Item item_{row_key:03d} (Excel Row {row_key + 1}): Missing file '{f}' (expected at '{fpath}').")

print('\nPrepared ' + str(finished) + ' records in \'' + options.items_dir + '\' folder.' )
