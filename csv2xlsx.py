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

    for csvfile in glob(glob_pattern):
        outdir = path.join(path.dirname(csvfile), outdir)
        makedirs(outdir, exist_ok=True)  # ensure outdir exists
        xlsxfile = path.join(outdir, path.basename(csvfile)[:-4] + ".xlsx")
        workbook = Workbook(xlsxfile)
        worksheet = workbook.add_worksheet()
        r, c = 0, -1
        with open(csvfile, "rt", encoding="utf8") as fp:

            for line in fp.readlines():
                line = line.strip("\n")
                for token in line.split(separator):
                    c += 1
                    token = token.strip(" \t")
                    if not token:
                        continue
                    if token.startswith('"'):
                        token = token.strip('"')
                        worksheet.write(r, c, token)
                    else:
                        if token == "nan":
                            continue
                        try:
                            worksheet.write_number(r, c, float(token))
                        except ValueError:  # Non-numbers that are not escaped with "
                            worksheet.write(r, c, token)
                c = -1
                r += 1
                print("\r" + next(spinner), end="")
        try:
            workbook.close()
            print(f"\rconverted: {csvfile} -> {xlsxfile}")
        except FileCreateError as e:
            print(f"\rfailed: {csvfile}\tyou probably have the file already open?")
            print(f"\t{e}")


if __name__ == "__main__":
    main()
