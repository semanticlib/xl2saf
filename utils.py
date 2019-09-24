import re

# Create index of header row
def parse_header(row):
  md_header = {}
  for idx, column in enumerate(row):
    # set defaults
    md_element = column.value.replace('dc.', '')
    md_qualifier = 'none'
    md_delimit = ''

    # Extract the delimiter
    if re.search(r'\(', md_element):
      (md_element, md_delimit) = re.sub(r'^([^\(]+)\s*\((.+)\)$', r'\1|\2', md_element).split('|')
    
    # Identify the qualifier
    if re.search(r'\.', md_element):
      (md_element, md_qualifier) = md_element.split(".")
    
    md_item = {
      'element': md_element,
      'qualifier': md_qualifier,
      'delimit': md_delimit
    }
    
    # prepare the index
    md_header[idx] = md_item
    idx+1

  return md_header

# Convert string to list
def get_list(text, delimiter):
  values = []
  if re.search(delimiter, text):
    values = text.split(delimiter)
  else:
    values.append(text)
  return values
