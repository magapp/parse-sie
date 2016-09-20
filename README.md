# parse-sie

This Python script can parse a [SIE file](http://www.sie.se/) . SIE is a Swedish standard file format for transfering economical data.

The data can either be exported as a CSV file that you can import into Excel. If you provide credentials to Google Sheet the data can be automatically imported to a Google Sheet.

When you have all transactions in Excel or Google Sheet, you can do great analysis and reports using Pivot tables and/or graphs. The reports generated from, for example SPCS or Fortnox are really embarrassing :rage:. The parser have support for block transactions, variables (kostnadsställen och projekt) and account names.

Syntax:

```
usage: parse-sie.py [-h] --filename FILENAME
                    [--encoding {cp850,latin-1,utf-8,windows-1252}]
                    [--output OUTPUT] [--debug] [--googlesheets GOOGLESHEETS]

Parse SIE file format

optional arguments:
  -h, --help            show this help message and exit
  --filename FILENAME
  --encoding {cp850,latin-1,utf-8,windows-1252}
  --output OUTPUT       Output CSV filename
  --debug               Display more debug
  --googlesheets GOOGLESHEETS
                        Name of Google Sheets to create. Need creds.json. Note
                        that the spreadsheet must exists and the content may
                        be overwritten.

```
