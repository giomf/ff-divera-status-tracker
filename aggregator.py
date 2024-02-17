import pandas
import os
import datetime
from pathlib import Path
import numpy
import openpyxl

_DATE_TIME_INPUT_FORMAT = '%d.%m.%Y %H:%M'
_DATE_TIME_OUTPUT_FORMAT = '%Y-%m-%dT%H:%M'

def get_exel_files(directory: Path) -> [Path]:
    return sorted(list(directory.glob('*.xlsx')))

def aggregate(directory: Path, output: Path):
    df = pandas.DataFrame()
    for idx, spread_sheet_path in enumerate(get_exel_files(directory)):
        wb = openpyxl.load_workbook(spread_sheet_path).active
        date = datetime.datetime.strptime(wb['B1'].value, _DATE_TIME_INPUT_FORMAT).strftime(_DATE_TIME_OUTPUT_FORMAT)
        status_sheet = pandas.read_excel(spread_sheet_path, skiprows=3, header=None, usecols=[0, 1], names=['Name', date], index_col='Name')
        df = pandas.concat([df, status_sheet], axis=1)

    df = df.transpose()
    df.to_excel(output)
