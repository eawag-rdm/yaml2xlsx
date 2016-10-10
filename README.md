# yaml2xlsx
Converts a file containing YAML-documents to an xlsx worksheet.

It is assumed that the YAML file contain one or several [*documents*](http://yaml.org/spec/1.2/spec.html#document//).
Each docuemnt is assumed to contains a [*mapping*](http://yaml.org/spec/1.2/spec.html#mapping//).
The top-row (header) of the worksheet contains all keys that appear in all documents.
Each *document* becomes one row in the worksheet.
Columns for which a document does not contain the respective key remain empty.


##Usage:

    yaml2xlsx [-h] INPUTFILE OUTPUTFILE

    Converts a file containing YAML-documents into an xlsx-worksheet. Each YAML-
    document represents one line.

    positional arguments:
      INPUTFILE   The YAML-file to convert
      OUTPUTFILE  The xlsx-file to write

    optional arguments:
      -h, --help  show this help message and exit

##Installation:

    pip install git+https://github.com/eawag-rdm/yaml2xlsx.git
