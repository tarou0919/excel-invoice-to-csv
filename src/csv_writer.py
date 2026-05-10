"""
複数の請求書データを1つのCSVに集約
"""
import csv
from pathlib import Path

from config import CSV_COLUMNS


def write_csv(records, output_path, encoding="utf-8-sig"):
    """
    抽出結果のリストをCSVに書き出す
    encoding: utf-8-sig はExcelで開いても文字化けしない
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with open(output_path, "w", newline="", encoding=encoding) as f:
        writer = csv.DictWriter(f, fieldnames=CSV_COLUMNS)
        writer.writeheader()
        for record in records:
            # 数値はカンマなしで、Noneは空文字に
            row = {}
            for col in CSV_COLUMNS:
                value = record.get(col)
                if value is None:
                    row[col] = ""
                elif isinstance(value, float) and value.is_integer():
                    row[col] = int(value)
                else:
                    row[col] = value
            writer.writerow(row)

    return output_path
