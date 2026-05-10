"""
Excel請求書から情報を抽出するコアモジュール
ラベル探索方式 + フォールバック処理
"""
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from datetime import datetime, date
import re
from pathlib import Path

from config import EXTRACTION_RULES, CSV_COLUMNS, HONORIFIC_KEYWORDS


def normalize_text(value):
    """セル値を文字列化して正規化(コロン・空白を除去)"""
    if value is None:
        return ""
    text = str(value).strip()
    # 末尾のコロン・全角コロンを除去
    text = re.sub(r"[::]\s*$", "", text)
    return text


def is_label_match(cell_text, labels):
    """セルテキストがラベル候補のいずれかと一致するかチェック"""
    normalized = normalize_text(cell_text)
    if not normalized:
        return False
    for label in labels:
        if normalized == label or normalized.startswith(label):
            return True
    return False


def get_neighbor_value(ws, row, col, direction, label_text=None):
    """
    指定セルの近隣から値を取得
    結合セルにも対応(結合の起点を取得)
    label_text が指定されれば、それと同じ値は採用しない(ラベル自身を返さないため)
    """
    candidates = []
    if direction == "right":
        candidates = [(row, col + i) for i in range(1, 6)]
    elif direction == "below":
        candidates = [(row + i, col) for i in range(1, 4)]
    elif direction == "right_or_below":
        candidates = [(row, col + i) for i in range(1, 6)] + \
                     [(row + i, col) for i in range(1, 4)]

    for r, c in candidates:
        try:
            cell = ws.cell(row=r, column=c)
            value = get_merged_cell_value(ws, cell)
            if value is None:
                continue
            value_str = str(value).strip()
            if not value_str:
                continue
            # ラベル自身と同じ値はスキップ(結合セル内を見たケース)
            if label_text and normalize_text(value) == normalize_text(label_text):
                continue
            return value
        except Exception:
            continue
    return None


def get_merged_cell_value(ws, cell):
    """結合セルの場合、結合範囲の起点セルの値を取得"""
    for merged_range in ws.merged_cells.ranges:
        if cell.coordinate in merged_range:
            top_left = ws.cell(row=merged_range.min_row, column=merged_range.min_col)
            return top_left.value
    return cell.value


def convert_value(value, data_type):
    """データ型に応じて値を変換"""
    if value is None:
        return None

    if data_type == "number":
        if isinstance(value, (int, float)):
            return float(value)
        # 文字列から数値抽出 "¥110,000" → 110000
        text = str(value).replace(",", "").replace("¥", "").replace("円", "").strip()
        try:
            return float(text)
        except ValueError:
            return None

    elif data_type == "date":
        if isinstance(value, (datetime, date)):
            return value.strftime("%Y-%m-%d")
        # 文字列の日付をパース
        text = str(value).strip()
        # よくある日本語形式: "2025年10月15日"
        m = re.match(r"(\d{4})\D+(\d{1,2})\D+(\d{1,2})", text)
        if m:
            y, mo, d = m.groups()
            try:
                return f"{int(y):04d}-{int(mo):02d}-{int(d):02d}"
            except ValueError:
                pass
        # ISO形式: "2025-10-15"
        m = re.match(r"(\d{4})-(\d{1,2})-(\d{1,2})", text)
        if m:
            y, mo, d = m.groups()
            return f"{int(y):04d}-{int(mo):02d}-{int(d):02d}"
        return text

    else:  # text
        text = str(value).strip()
        # 改行を空白に
        text = re.sub(r"\s+", " ", text)
        return text


def get_merged_range_for_cell(ws, cell):
    """セルが結合セルの一部なら、その結合範囲を返す。違えば None"""
    for merged_range in ws.merged_cells.ranges:
        if cell.coordinate in merged_range:
            return merged_range
    return None


def find_value_by_labels(ws, rule):
    """
    シート内をラベル探索して値を取得
    結合セルがラベルになっている場合、結合範囲の外側を探す
    """
    labels = rule["labels"]
    direction = rule["direction"]

    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            if not is_label_match(cell.value, labels):
                continue

            label_text = str(cell.value)
            # 結合セルの場合、結合範囲の右端/下端から探索開始
            merged = get_merged_range_for_cell(ws, cell)
            if merged:
                start_row = merged.max_row
                start_col = merged.max_col
            else:
                start_row = cell.row
                start_col = cell.column

            value = get_neighbor_value(ws, start_row, start_col, direction, label_text)
            if value is not None:
                return value
    return None


def find_customer_by_honorific(ws):
    """
    取引先名のフォールバック: 「御中」「様」を含むセルを探す
    """
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            text = str(cell.value).strip()
            for keyword in HONORIFIC_KEYWORDS:
                if keyword in text:
                    # 「株式会社サンプル商事 御中」のようなセル全体を返す
                    return text
    return None


def find_total_by_max_amount(ws):
    """
    合計金額のフォールバック: 数値セルの中で最大値を取得
    (請求書では合計金額が最大数値であることが多い)
    """
    max_val = None
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, (int, float)) and cell.value > 0:
                if max_val is None or cell.value > max_val:
                    max_val = cell.value
    return max_val


def extract_from_file(file_path):
    """
    1つのExcelファイルから請求書情報を抽出して dict で返す
    """
    file_path = Path(file_path)
    result = {col: None for col in CSV_COLUMNS}
    result["ファイル名"] = file_path.name

    try:
        wb = load_workbook(file_path, data_only=True)
    except Exception as e:
        print(f"  ⚠️  ファイル読込失敗: {file_path.name} ({e})")
        return result

    # 全シートを順番に走査(最初のシートを優先)
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for field, rule in EXTRACTION_RULES.items():
            if result.get(field) is not None:
                continue  # 既に取得済みの場合はスキップ
            value = find_value_by_labels(ws, rule)
            if value is not None:
                converted = convert_value(value, rule["data_type"])
                if converted is not None and converted != "":
                    result[field] = converted

    # フォールバック処理
    ws_main = wb[wb.sheetnames[0]]

    if not result.get("取引先名"):
        customer = find_customer_by_honorific(ws_main)
        if customer:
            result["取引先名"] = customer

    if not result.get("合計金額"):
        total = find_total_by_max_amount(ws_main)
        if total:
            result["合計金額"] = total

    # 整合性チェック: 小計+消費税 ≒ 合計金額
    if (result.get("小計") and result.get("消費税") and not result.get("合計金額")):
        result["合計金額"] = result["小計"] + result["消費税"]

    return result


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        path = sys.argv[1]
        data = extract_from_file(path)
        print(f"\n=== {path} ===")
        for k, v in data.items():
            print(f"  {k}: {v}")
