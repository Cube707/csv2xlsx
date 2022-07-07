# csv2xlsx

A utility for converting English (US) style CSV files into German (EU) style Excel worksheets.

It replaces `,` as a column separator by `;` and converts the `.` as a decimal separator to `,`


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

You can pass it a [Glob-pattern](https://www.malikbrowne.com/blog/a-beginners-guide-glob-patterns) to be used to find the files. By default, `*.csv` is
used.

You can use the `--outdir` option to specify an output directory, either relative to the current folder or as an absolute path.


## Documentation

You can access this documentation by passing the `--help` option to the script.

```
Usage: csv2xlsx.py [OPTIONS] [GLOB_PATTERN]

  A utility for convernverting englisch (US) style CSV files into German (EU)     
  style Excel worksheets. It replaces ',' as a column seperator by ';' and        
  converts the . as a decimal separator to ','

Options:
  -h, --help              Show this message and exit.
  -o, --outdir DIRECTORY  optional output directory
```

# build the binary

by using pyinstaller:

```
pyinstaller -F -p .venv/Lib/site-packages csv2xlsx.py
```
