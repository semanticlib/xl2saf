import re
from config import LOG_FILE
import logging

logging.basicConfig(filename=LOG_FILE, level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def parse_header(row):
    """
    Parses the header row of the Excel file to extract metadata field information.
    """
    md_header = {}
    for idx, column in enumerate(row):
        header_cell_value = str(column.value) if column.value is not None else ""
        md_element, md_qualifier, md_delimit = _parse_header_cell(header_cell_value)
        
        md_item = {
            'element': md_element,
            'qualifier': md_qualifier,
            'delimit': md_delimit
        }
        md_header[idx] = md_item
    return md_header

def _parse_header_cell(header_cell_value):
    """
    Parses a single header cell to extract element, qualifier, and delimiter.
    """
    md_element = header_cell_value.replace('dc.', '')
    md_qualifier = 'none'
    md_delimit = ''

    # Extract the delimiter
    match = re.match(r'^([^\(]+)\s*\((.+)\)
, md_element)
    if match:
        md_element = match.group(1).strip()
        md_delimit = match.group(2).strip()

    # Identify the qualifier
    if '.' in md_element:
        parts = md_element.split(".", 1)
        md_element = parts[0]
        if len(parts) > 1:
            md_qualifier = parts[1]

    return md_element.strip(), md_qualifier.strip(), md_delimit

def get_list(text, delimiter):
    """
    Converts a delimited string to a list of strings.
    """
    if text is None:
        return []
    text_str = str(text)
    if not text_str.strip():
        return []

    if not delimiter:
        return [text_str.strip()]
    
    return [part.strip() for part in text_str.split(delimiter) if part.strip()]
