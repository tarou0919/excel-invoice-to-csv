"""
複数の請求書データを1つのCSVに集約
"""
import csv
import logging
from pathlib import Path
from typing import List, Dict, Any

logger = logging.getLogger(__name__)


def write_csv(
    records: List[Dict[str, Any]],
    output_path,
    csv_columns: List[str],
    encoding: str = "utf-8-sig",
) -> Path:
    """
    抽出結果のリストをCSVに書き出す

    Args:
        records: 抽出結果の辞書リスト
        output_path: 出力先CSVパス
        csv_columns: 出力する列の順序
        encoding: 文字コード(utf-8-sig はExcelで開いても文字化けしない)

    Returns:
        実際に書き出したCSVファイルのパス
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        with open(output_path, "w", newline="", encoding=encoding) as f:
            writer = csv.DictWriter(f, fieldnames=csv_columns, extrasaction="ignore")
            writer.writeheader()
            for record in records:
                row = {}
                for col in csv_columns:
                    value = record.get(col)
                    if value is None:
                        row[col] = ""
                    elif isinstance(value, float) and value.is_integer():
                        row[col] = int(value)
                    else:
                        row[col] = value
                writer.writerow(row)
    except PermissionError:
        logger.error(
            f"CSV書き出し失敗: '{output_path}' が他のアプリで開かれている可能性があります"
        )
        raise
    except Exception as e:
        logger.error(f"CSV書き出し失敗: {type(e).__name__}: {e}")
        raise

    return output_path
