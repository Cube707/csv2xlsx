from glob import glob
from itertools import cycle
from os import path, makedirs

import click
from xlsxwriter.workbook import Workbook


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
    It converts the . as a decimal separator to ','. Tokens wrapped in "" are treated as strings and not converted.

    All files identified by GLOB_PATTERN are converted one by one.
    """

    for csvfile in glob(glob_pattern):
        outdir = path.join(path.dirname(csvfile), outdir)
        makedirs(outdir, exist_ok=True)  # ensure outdir exists
        xlsxfile = path.join(outdir, path.basename(csvfile)[:-4] + ".xlsx")
        workbook = Workbook(xlsxfile)
        worksheet = workbook.add_worksheet()
        r, c = 0, 0
        with open(csvfile, "rt", encoding="utf8") as fp:

            for line in fp.readlines():
                for token in line.split(separator):
                    if not token.startswith('"'):
                        token = token.replace(".", ",")
                    worksheet.write(r, c, token)
                    c += 1
                c = 0
                r += 1
                print("\r" + next(spinner), end="")
        workbook.close()
        print(f"\rconverted: {csvfile} -> {xlsxfile}")


if __name__ == "__main__":
    spinner = cycle(["-", "/", "|", "\\"])
    main()
