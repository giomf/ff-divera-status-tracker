import pandas as pd
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
    result = pd.DataFrame()
    for idx, status_part in enumerate(get_exel_files(directory)):
        wb = openpyxl.load_workbook(status_part).active
        date = datetime.datetime.strptime(wb['B1'].value, _DATE_TIME_INPUT_FORMAT).strftime(_DATE_TIME_OUTPUT_FORMAT)
        status_sheet = pd.read_excel(status_part, skiprows=3, header=None, usecols=[0, 1], names=['Name', date], index_col='Name')
        result = pd.concat([result, status_sheet], axis=1)

    result = result.transpose()
    result.to_excel(output)
