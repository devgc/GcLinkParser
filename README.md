# GcLinkParser
A GC link parser for both linkfiles and jumplists.

Note: For best results use the -d options to process a directory due to encoding issues and the CMD.

# Usage
```
usage: GcLinkParser.py [-h] [-f FILE_NAME] [-d DIRECTORY] [--pipe] [--jmp]
                       [--timeformat TIMEFORMAT] [--timezone TIMEZONE]
                       [--listtz] [--txt] [--sqlite SQLITE_DB]
                       [--delimiter DELIMITER] [--json] [--eshost ESHOST]
                       [--index INDEX]

GcLinkParser v1.00 [Copywrite G-C Partners, LLC 2015,2016]

EXAMPLES:
========================================================================
List Supported Timezones
GcLinkParser.exe --tzlist
------------------------------------------------------------------------
JSON Output
GcLinkParser.exe -f LINKFILE --json
------------------------------------------------------------------------
CSV Output
GcLinkParser.exe -f LINKFILE --txt
------------------------------------------------------------------------
Send records to Elasticsearch
GcLinkParser.exe -f LINKFILE --eshost "ELASTIC_IP" --index lnkfiles
------------------------------------------------------------------------
Get Filelist from dir and format txt
dir /b /s /a *.lnk | GcLinkParser.exe --pipe --txt

NOTES:
========================================================================
The AppId is enumerated from the list found at
http://forensicswiki.org/wiki/List_of_Jump_List_IDs
that was last modified as of 4 March 2015, at 11:07.

You can create a custom AppId list and stick it in the cwd of this tool.
Name the file 'AppIdList.txt' and should be formated as 16HEXID\tAPP_NAME

optional arguments:
  -h, --help            show this help message and exit
  -f FILE_NAME, --file FILE_NAME
                        lnk filename
  -d DIRECTORY, --directory DIRECTORY
                        directory that contains linkfiles or jumplists
  --pipe                get filelist from pipe (dir /b /s /a *.lnk)
  --jmp                 Parse files as jumplists
  --timeformat TIMEFORMAT
                        datetime format
  --timezone TIMEZONE   output timezone
  --listtz              list all supported timezone options
  --txt                 output to text file (default delimiter is \t)
                        [Recommended not to use ',' as delimiter]
  --sqlite SQLITE_DB    output to a specified sqlite file
  --delimiter DELIMITER
                        csv delimiter
  --json                json output
  --eshost ESHOST       Elastic host
  --index INDEX         Elastic index
```

# Dependencies
You will need the following libraries:
- elastichandler
  - Project Page:</br>
  https://github.com/devgc/ElasticHandler
- liblnk
  - Project Page:</br> 
  https://github.com/libyal/liblnk
  - Compiled binary:</br>
  https://github.com/log2timeline/l2tbinaries
- libfwsi
  - Project Page:</br>
  https://github.com/libyal/libfwsi
  - Compiled binary:</br>
  https://github.com/log2timeline/l2tbinaries
- libolecf
  - Project Page:</br>
  https://github.com/libyal/libolecf
  - Compiled binary:</br>
  https://github.com/log2timeline/l2tbinaries

Special Thanks to Joachim Metz! His libraries make this work.

See licenses folder for library licenses.
