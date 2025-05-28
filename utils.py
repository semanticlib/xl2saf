import re

# Create index of header row
def parse_header(row):
  md_header = {}
  for idx, column in enumerate(row):
    header_cell_value = str(column.value) if column.value is not None else ""

    # set defaults
    md_element = header_cell_value.replace('dc.', '')
    md_qualifier = 'none'
    md_delimit = ''

    # Extract the delimiter
    if re.search(r'\(', md_element):
      match = re.match(r'^([^\(]+)\s*\((.+)\)$', md_element)
      if match:
          md_element = match.group(1).strip()
          md_delimit = match.group(2).strip()
    
    # Identify the qualifier
    if '.' in md_element:
      parts = md_element.split(".", 1)
      md_element = parts[0]
      if len(parts) > 1:
          md_qualifier = parts[1]
    
    md_item = {
      'element': md_element.strip(),
      'qualifier': md_qualifier.strip(),
      'delimit': md_delimit # Already stripped if found, or empty
    }
    
    # prepare the index
    md_header[idx] = md_item

  return md_header

# Convert string to list, stripping whitespace from elements and removing empty elements
def get_list(text, delimiter):
  if text is None:
      return []
  text_str = str(text)
  if not text_str.strip(): # If text is empty or only whitespace
      return []

  if not delimiter: # If delimiter is empty string or None, treat as single value
      return [text_str.strip()]
  
  values = [part.strip() for part in text_str.split(delimiter) if part.strip()]
  return values
