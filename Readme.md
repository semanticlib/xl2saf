# XL2SAF

[![Test Status](https://github.com/semanticlib/xl2saf/actions/workflows/test.yml/badge.svg)](https://github.com/semanticlib/xl2saf/actions/workflows/test.yml)

**Excel to DSpace [Simple Archive Format (SAF)](https://wiki.lyrasis.org/pages/viewpage.action?pageId=104566653) conversion script**

#### Dependencies
This script requires Python 3. All dependencies are listed in the `requirements.txt` file.

## Prepare Source XLSX/Open Office XML (.xlsx)
The source files must be in [XLSX/Open Office XML](https://www.loc.gov/preservation/digital/formats/fdd/fdd000398.shtml) format. Most spreadsheet applications, including MS Excel, OpenOffice/LibreOffice Calc and Google Sheets, support export to XLSX format.

The first row must be header row in the format `dc.element_name[.qualifier][(separator)]`, where the `qualifier` and `separator` are optional. e.g.
```
dc.title
dc.title.alternative
dc.contributor.author
```

If the field data is multi-valued, it must use a separator that is not part of the data, such as '|' or ';'. The same must be specified in the header, e.g.
```
dc.contributor.author(;)
```

Populate the `filenames(;)` column with the full-text files. File names must be relative to the `--base_dir` option. e.g.
```
Article1.pdf; Slides1.ppt
```

See [Metadata and Bitstream Format Registries](https://wiki.lyrasis.org/display/DSDOC7x/Metadata+and+Bitstream+Format+Registries) page for available metadata fields.

Check the [sample input file](./sample_data/Input.xlsx) as an example. Prepare separate input file for each collection.

## Installation / Setup
```bash
git clone https://github.com/semanticlib/xl2saf.git
cd xl2saf

# Install dependencies globally (if having sudo access)
sudo pip3 install -r requirements.txt

# Alternatively, install dependencies in virtual environment
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

## Usage
```shell
python xl2saf.py -h
```
```
Usage: xl2saf.py [options]

options:
  -h, --help            show this help message and exit
  -f, --input_file XLSX_FILE
                        Input XLSX file
  -b, --base_dir BASE_DIR
                        Base directory for full-text files
  -s, --start START     Start processing from this Excel row (default: 2)
  -e, --end END         Stop processing at this Excel row (inclusive)
  -d, --items_dir ITEMS_DIR
                        Output directory for converted files (default: import)
  --log_file LOG_FILE   Log file path (default: xl2saf.log)

```

#### Example with sample data:
```bash
python xl2saf.py -f sample_data/Input.xlsx -b sample_data/fulltext
```
Or
```bash
python xl2saf.py --input_file=sample_data/Input.xlsx --base_dir=sample_data/fulltext --items_dir=items_import --start=2 --end=10
```

The above comman should be run for each collection using separate `--input_file` and `--items_dir`.

## Importing into DSpace
See detailed instructions in [DSpace Wiki](https://wiki.lyrasis.org/pages/viewpage.action?pageId=104566653):

**Using web interface**
Compress the output folder into a Zip file to import them using the DSpace web interface. This method is suitable for small datasets.

**Import from command line**
Use the command-line tool for large datasets.

Example command:
```bash
cd /path/to/dspace/backend
bin/dspace import --add --eperson=joe@user.com --collection=1/12 --source=items_import --mapfile=mapfile
```

## Testing

To ensure the script is working correctly, you can run the included tests. First, install the dependencies. Then, run the tests using the following command:

```bash
python -m unittest discover
```

Create an issue to report any problem.
