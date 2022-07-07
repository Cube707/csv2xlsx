from glob import glob
from os import path

import click
from xlsxwriter.workbook import Workbook


@click.command()
@click.help_option("--help", "-h")
@click.option(
    "--outdir",
    "-o",
    type=click.Path(exists=True, file_okay=False),
    help="optional output directory",
    default=".",
)
@click.argument("glob_pattern", default="*.csv", nargs=1)
def main(glob_pattern, outdir):
    """A utility for converting English (US) style CSV files into German (EU) style Excel worksheets.
    It replaces ',' as a column separator by ';' and converts the . as a decimal separator to ','
    """

    for csvfile in glob(glob_pattern):
        xlsxfile = csvfile[:-4] + ".xlsx"
        workbook = Workbook(path.join(outdir, xlsxfile))
        worksheet = workbook.add_worksheet()
        r, c = 0, 0
        with open(csvfile, "rt", encoding="utf8") as fp:

            for line in fp.readlines():
                for token in line.split(","):
                    if not token.startswith('"'):
                        token = token.replace(".", ",")
                    worksheet.write(r, c, token)
                    c += 1
                c = 0
                r += 1
        workbook.close()
        print(f"converted: {csvfile} -> {xlsxfile}")


if __name__ == "__main__":
    main()
