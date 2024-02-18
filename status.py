import pandas as pd
from pathlib import Path

def calculate_on_duty_percentage(input: Path, off_duty_keyword: str, output: Path):
    result = pd.DataFrame()
    data = pd.read_excel(input, index_col=0)
    for name, series in data.items():
        percentage = _calculate_on_duty_percentage(series, off_duty_keyword)
        result[name] = [percentage]

    result.transpose().to_excel(output, header=None)

def _calculate_on_duty_percentage(data: pd.Series, off_duty_keyword: str) -> int:
    total = data.count()
    off_duty_count = data.str.count(off_duty_keyword).sum()
    return '%.1f' % (100 - (off_duty_count / total * 100))

