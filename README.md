[![GitHub badge](https://badges.aleen42.com/src/github_flat_square.svg)](https://github.com/Cube707/csv2xlsx)
[![Python badge](https://badges.aleen42.com/src/python_flat_square.svg)](https://python.org)
[![Terminal badge](https://badges.aleen42.com/src/cli_flat_square.svg)](https://www.techtarget.com/searchwindowsserver/definition/command-line-interface-CLI)

# csv2xlsx

A utility for converting English (US) style CSV files into German (EU) style Excel worksheets.

It converts the `.` as a decimal separator to `,`. Tokens wrapped in `"` are treated as strings and not converted, but get their `"` striped.


## install dependencies

You only need to install dependencies for the python script. The prebuild binary comes bundled with everything!

A virtual environment is recommended, but optional:

1. Set up the venv:

    ```
    python -m venv .venv
    ```

2. Activate it:

    ```
    .venv\Scripts\activate
    ```

3. install the dependencies:

    ```
    pip install -r requirements.txt
    ```

Done!


## Usage

use as a python script:

```
python csv2xlsx.py 
```

use the prebuild binary by simply running it:

```
csv2xlsx.exe 
```

Without any options this will format all `.csv` files in the current working directory.

Additionally you can pass the following options to the scripts to modify its behaviour. The call signature looks like this:

```bash
Usage: csv2xlsx.py [OPTIONS] [GLOB_PATTERN]
```

### `GLOB_PATTERN` positional argument

You can pass a glob-pattern to the script to tell it what files to modify. By default, `*.csv` is used, but you can modify it as you need.

Here are some examples and results:

```bash
# same as with default settings:
csv2xlsx.py *.csv

> converted: file.csv -> file.xlsx
```

```bash
# recurse through all subdirectorys as well:
csv2xlsx.py **/*.csv

> converted: file.csv -> file.xlsx
> converted: sub/file.csv -> sub/file.xlsx
```

```bash
# using an absoloute path:
csv2xlsx.py /absoloute/path/*.csv

> converted: /absoloute/path/file.csv -> /absoloute/path/file.xlsx
```

### `--outdir` / `-o` option

You can use the `--outdir` option to specify an output directory, either relative to the current folder or as an absolute path. Directories will be
created if required

Here are some examples and results:

```bash
# relative output directory:
csv2xlsx.py -o out

> converted: file.csv -> out/file.xlsx
```

```bash
# recurse through all subdirectories as well:
csv2xlsx.py -o out **/*.csv

> converted: file.csv -> out/file.xlsx
> converted: sub/file.csv -> sub/out/file.xlsx
```

```bash
# use an absoloute output directory:
csv2xlsx.py -o /absoloute/output/path/ /absoloute/path/*.csv

> converted: /absoloute/path/file.csv -> /absoloute/output/path/file.xlsx
```

### `--separator` / `-s` option

Defines the separator used in the CSV files. Defaults to `,`, but can be set to any string.

Here is a examples:

```bash
csv2xlsx.py -s ";" *.csv
```

### `--help` / `-h` option

Displays a help text with a short version of this documentation.


# build the binary

by using pyinstaller:

```
pyinstaller -F -p .venv/Lib/site-packages csv2xlsx.py
```

This will produce the windows binary in the `dist/` subfolder.
