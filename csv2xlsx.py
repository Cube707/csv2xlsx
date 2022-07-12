from glob import glob
from itertools import cycle
from os import path, makedirs

import click
from xlsxwriter.workbook import Workbook, FileCreateError


@click.command()
@click.help_option("--help", "-h")
@click.option(
    "--outdir",
    "-o",
    type=click.Path(file_okay=False),
    help="optional output directory",
    default=".",
)
@click.option(
    "--separator",
    "-s",
    type=str,
    help="separator of the CSV file, defaults to ','",
    default=",",
)
@click.argument("glob_pattern", default="*.csv", nargs=1)
def main(glob_pattern, outdir, separator):
    """A utility for converting English (US) style CSV files into German (EU) style Excel worksheets.
    It converts the . as a decimal separator to ','. Tokens wrapped in "" are treated as strings and not converted,
    but get their " striped.

    All files identified by GLOB_PATTERN are converted one by one.
    """

    spinner = cycle(["-", "/", "|", "\\"])  # a spinner to dislay "work in progress"

    # loop over all files matching the glob pattern
    for csvfile in glob(glob_pattern):

        # handle output dir:
        outdir = path.join(path.dirname(csvfile), outdir)
        makedirs(outdir, exist_ok=True)  # ensure outdir exists

        # create workbook object:
        xlsxfile = path.join(outdir, path.basename(csvfile)[:-4] + ".xlsx")
        workbook = Workbook(xlsxfile)
        worksheet = workbook.add_worksheet()

        r = 0
        c = -1  # we start at c=-1 because the first action is to increment c to 0
        with open(csvfile, "r") as fp:

            for line in fp.readlines():
                # remove newline characters from the end of the line:
                line = line.strip("\n")

                # loop over all the elements in a line:
                for token in line.split(separator):
                    c += 1  # first thing we do is increment so we can skip empty tokens
                    token = token.strip(" \t")  # remove uneeded whitespace
                    if not token:
                        continue  # skip empty tokens

                    # escaped "string" tokens are written as is:
                    if token.startswith('"'):
                        token = token.strip('"')
                        worksheet.write(r, c, token)

                    # all not-strings are assumed to be numbers:
                    else:
                        if token == "nan":
                            continue  # skip nan's
                        try:
                            worksheet.write_number(r, c, float(token))
                        except ValueError:  # Non-numbers that are not escaped with "
                            worksheet.write(r, c, token)
                r += 1
                c = -1
                print("\r" + next(spinner), end="")  # display "work in progress"
        try:
            workbook.close()
            print(f"\rconverted: {csvfile} -> {xlsxfile}")
        except FileCreateError as e:
            print(f"\rfailed: {csvfile}\tyou probably have the file already open?")
            print(f"\t{e}")


if __name__ == "__main__":
    main()
