# XL2SAF

[![Test Status](https://github.com/semanticlib/xl2saf/actions/workflows/test.yml/badge.svg)](https://github.com/semanticlib/xl2saf/actions/workflows/test.yml)

**Excel to DSpace [Simple Archive Format (SAF)](https://wiki.lyrasis.org/display/DSDOC7x/Importing+and+Exporting+Items+via+Simple+Archive+Format) conversion script**

#### Dependencies
Python3 with following dependencies
- openpyxl
- lxml

**Use pip to install dependencies**

On Debian/Ubuntu, following should work:
```
$ sudo apt install python3-pip
$ pip3 install openpyxl lxml
```

On Windows, pip3 should already be there. Just install the dependencies:
```
C:\> pip3 install openpyxl lxml
```


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

## Testing

To ensure the script is working correctly, you can run the included tests. First, install the dependencies:

```
$ pip install -r requirements.txt
```

Then, run the tests using the following command:

```
$ python -m unittest discover
```

## Convert to SAF

**Clone this repository**
```
$ git clone https://github.com/semanticlib/xl2saf.git
$ cd xl2saf
```

**Run the conversion script**
```
$ python xl2saf.py -h
Usage: xl2saf.py [options]

Options:
  -h, --help            show this help message and exit
  -f FILE, --input_file=FILE
                        Input XLSX file
  -b BASE_DIR, --base_dir=BASE_DIR
                        Base directory for full-text files
  -d ITEMS_DIR, --items_dir=ITEMS_DIR
                        Output directory for converted files

```

Try this command for sample data:
```
$ python xl2saf.py -f sample_data/Input.xlsx -b sample_data/fulltext -d collection_12
```
OR
```
$ python xl2saf.py --input_file=sample_data/Input.xlsx --base_dir=sample_data/fulltext --items_dir=collection_12
```

The above comman should be run for each collection using separate `--input_file` and `--items_dir`.

## Importing into DSpace
See detailed instructions in [DSpace Wiki](https://wiki.lyrasis.org/display/DSDOC7x/Importing+and+Exporting+Items+via+Simple+Archive+Format):

**Using web interface**
Compress the output folder into a Zip file to import them using the DSpace web interface. This method is suitable for small datasets.

**Import from command line**
Use the command-line tool for large datasets.

Example command:
```
$ [dspace_dir]/bin/dspace import --add --eperson=joe@user.com --collection=1/12 --source=collection_12 --mapfile=mapfile
```

Create an issue to report any problem.
