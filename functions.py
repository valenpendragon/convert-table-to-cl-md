import os
from pathlib import Path
import pandas as pd

def get_table(filepath, table_name=""):
    """This function takes a filepath to a spreadsheet file and the table
    or tab name of the desired spreadsheet. It returns either a string, which
    is an error message, or a pandas DataFrame taken from the file. Valid file
    extensions are csv, xls, and xlsx. table_name is only required for Excel
    files. pandas.read_csv cannot take sheet_name as an argument.
    :param filepath: path to the csv or Excel file
    :param table_name: this is the worksheet name as it appears on the tab
    :return: str (error message) or pandas.DataFrame
    """
    extension = Path(filepath).suffix
    filename = Path(filepath).name
    match extension:
        case "csv":
            try:
                df = pd.read_csv(filepath, na_filter=False)
            except UnicodeError:
                return "File format is not CSV"
            except Exception as err:
                return f"Unexpected {err=}, {type(err)=}."
            else:
                return df
        case "xls" | "xlsx":
            try:
                df = pd.read_excel(filepath, sheet_name=table_name,
                                   na_filter=False)
            except ValueError:
                return f"Worksheet {table_name} is not in {filename}."
            except Exception as err:
                return f"Unexpected {err=}, {type(err)=}"
            else:
                return df

        case _:
            return f"{filename} has invalid extension {extension}."

def write_output(dest_dir, output_file, output_lines):
    """This function writes the converted rows as lines of text in the specified
    output_file. If the file extension is not specified, it will use txt. The
    This program alwayws overwrites the target file.
    :param dest_dir: folder to write the new file in
    :param output_file: name of the file to use
    :param output_lines: list of text lines to write to the target file
    :return:
    """
    extension = Path(output_file).suffix
    if extension == "":
        extension = "txt"
        targetfile = f"{output_file}.{extension}"
    filepath = f"{dest_dir}/{targetfile}"
    with open(filepath, "w") as f:
        for line in output_lines:
            f.writelines(line)