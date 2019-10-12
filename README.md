# mcte
Copy multipule csv files to excel file. Create new excel file based on the template.xlsx.

## Install
```
git clone https://github.com/yuitoku/mcte.git
pip install -r requirements.txt
```

## Usage
Please store csv files to csv directory. Then exec mcte.py.
```
python mcte.py
```

Following is deteiled usage. If change layout of sheet, modify template.xlsx.
```
usage: python mcte.py [--help] [--file <new file name> ] [--row <number] [--column <number>]
                      [--delimiter <delimiter>] [--font <font style>] [--size <font size> ]

Copy multipule csv files to excel.
Required template.xlsx in template directory, and csv files in csv directories.

optional arguments:
  -h, --help            show this help message and exit
  -f <new file name>, --file <new file name>
                        file name for destination workbook
  -r <number>, --row <number>
                        destination row number
  -c <number>, --column <number>
                        destination column number
  -d <delimiter>, --delimiter <delimiter>
                        delimiter from csv file name to destination sheet name
  --font <font style>   destination font style
  --size <font size>    destination font size
  ```

## Licence
[MIT License](https://github.com/yuitoku/mcte/blob/master/LICENSE.txt) @ Yuitoku