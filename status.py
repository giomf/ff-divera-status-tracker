import pandas as pd
from pathlib import Path

_HEADER=['Total on duty', 'Weekend on duty', 'Off duty without note']

def calculate_on_duty_percentage(input: Path, off_duty_keyword: str, output: Path):
    result = pd.DataFrame()
    data = pd.read_excel(input, index_col=[0, 1], parse_dates=['Date'])
    names = data.index.get_level_values(0).drop_duplicates()

    for name in names:
        name_data = data.loc[name]
        status = name_data['Status']
        total_percentage = _calculate_on_duty_total_percentage(status, off_duty_keyword)
        weekend_percentage = _calculate_on_duty_weekend_percentage(status, off_duty_keyword)
        without_note_percentage = _calculate_off_duty_without_note_percentage(name_data, off_duty_keyword)
        result[name] = [total_percentage, weekend_percentage, without_note_percentage]

    result.transpose().to_excel(output, float_format='%.1f', index_label='Name', header=_HEADER)

def _calculate_on_duty_total_percentage(data: pd.Series, off_duty_keyword: str) -> float:
    total_count = data.count()
    total_off_duty_count = data.str.count(off_duty_keyword).sum()

    return _calculate_on_duty_percentage_from_off_duty(total_off_duty_count, total_count)

def _calculate_on_duty_weekend_percentage(data: pd.Series, off_duty_keyword: str) -> float:
    weekend = data[data.index.dayofweek > 4]
    weekend_total_count = weekend.count()
    weekend_off_duty_count = weekend.str.count(off_duty_keyword).sum()

    return _calculate_on_duty_percentage_from_off_duty(weekend_off_duty_count, weekend_total_count)

def _calculate_off_duty_without_note_percentage(data: pd.DataFrame, off_duty_keyword: str) -> float:
    total_off_duty_count = data['Status'].str.count(off_duty_keyword).sum()   
    off_duty_without_note = len(data[(data['Status'] == off_duty_keyword) & (data['Note'].isnull())])    
    return _calculate_percentage(off_duty_without_note, total_off_duty_count)

def _calculate_on_duty_percentage_from_off_duty(part: int, total: int) -> float:
    return 100 - _calculate_percentage(part, total)

def _calculate_percentage(part: int, total: int) -> float:
    return part / total * 100
