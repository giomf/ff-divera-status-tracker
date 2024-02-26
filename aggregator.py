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
    for _, status_part in enumerate(get_exel_files(directory)):
        wb = openpyxl.load_workbook(status_part).active
        date = datetime.datetime.strptime(wb['B1'].value, _DATE_TIME_INPUT_FORMAT).strftime(_DATE_TIME_OUTPUT_FORMAT)
        status_sheet = pd.read_excel(status_part, skiprows=3, header=None, usecols=[0, 1, 3], names=['Name', 'Status', 'Note'], index_col='Name')

        index = pd.MultiIndex.from_product([status_sheet.index, [date]], names=['Name', 'Date'])
        data = {'Status': status_sheet['Status'].values, 'Note': status_sheet['Note'].values}
        status_at_date = pd.DataFrame(data, index=index)

        result = pd.concat([result, status_at_date])
    result.sort_index().to_excel(output)
