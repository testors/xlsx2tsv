# xlsx2tsv
xlsx 파일을 tsv 파일로 고속으로 변환해 줌

```
Usage: ./xlsx_to_tsv <input.xlsx> [start_row] [--no-wildcard]
  start_row: 1-based row number to start conversion (default: 1)
  Note: Output files will be named as '<SheetName>.tsv'
        Only sheets with names containing A-Z, a-z, 0-9, -, _, * will be processed
        Sheets with spaces, special characters, or non-ASCII will be skipped
        * characters in sheet names will be removed from output filenames
        --no-wildcard: Do not allow * characters in sheet names
```
