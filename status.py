import pandas as pd
from pathlib import Path

def calculate_on_duty_percentage(input: Path, off_duty_keyword: str, output: Path):
    result = pd.DataFrame()
    data = pd.read_excel(input, index_col=0, parse_dates=True)

    for name, series in data.items():
        total_percentage = _calculate_on_duty_total_percentage(series, off_duty_keyword)
        weekend_percentage = _calculate_on_duty_weekend_percentage(series, off_duty_keyword)
        result[name] = [total_percentage, weekend_percentage]

    result.transpose().to_excel(output, float_format='%.1f', index_label='Name', header=['Total on duty', 'Weekend on duty'])

def _calculate_on_duty_total_percentage(data: pd.Series, off_duty_keyword: str) -> float:
    total_count = data.count()
    total_off_duty_count = data.str.count(off_duty_keyword).sum()

    return _calculate_on_duty_percentage_from_off_duty(total_off_duty_count, total_count)

def _calculate_on_duty_weekend_percentage(data: pd.Series, off_duty_keyword: str) -> float:
    weekend = data[data.index.dayofweek > 4]
    weekend_total_count = weekend.count()
    weekend_off_duty_count = weekend.str.count(off_duty_keyword).sum()

    return _calculate_on_duty_percentage_from_off_duty(weekend_off_duty_count, weekend_total_count)

def _calculate_on_duty_percentage_from_off_duty(part: int, total: int) -> float:
    return 100 - (part / total * 100)
