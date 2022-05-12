# xlsx-to-csv

Convert XLSX spreadsheets to CSVs. Built with [Excelize](https://github.com/xuri/excelize).

## Usage

```console
$ xlsx-to-csv --help                                                                                                                                                                                        main
Usage:
  xlsx-to-csv [OPTIONS]

Application Options:
  -l, --list-sheets  List sheets
  -s, --sheet=       Sheet to convert
  -i, --input=       Input XLSX file (default: stdin)
  -o, --output=      Output CSV file (default: stdout)

Help Options:
  -h, --help         Show this help message
```

### List sheets

```console
$ xlsx-to-csv --list-sheets -i data.xlsx
Downloads
Cumulative
```

### Convert sheet to CSV

```console
$ xlsx-to-csv --sheet Downloads -i data.xlsx
Date,Downloads
05-12-22,918
05-13-22,681
05-14-22,360
05-15-22,372
05-16-22,255
05-17-22,152
05-18-22,875
```
