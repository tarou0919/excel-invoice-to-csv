"""
抽出ルール設定
ラベル(キーワード)を見つけて、その近隣セルから値を取得する方式
"""

# CSVに出力する項目の定義
# 各項目について、以下を指定:
#   - labels: Excel上で探すラベル候補(複数指定可)
#   - direction: 値を取得する方向 ("right" or "below")
#   - sheet_priority: 検索するシートの優先順(空なら全シート)
#   - data_type: "text", "number", "date" のいずれか

EXTRACTION_RULES = {
    "請求書番号": {
        "labels": ["請求書番号", "Invoice No", "No.", "請求No"],
        "direction": "right",
        "data_type": "text",
    },
    "発行日": {
        "labels": ["発行日", "請求日", "Issue Date", "Date"],
        "direction": "right",
        "data_type": "date",
    },
    "支払期限": {
        "labels": ["支払期限", "支払期日", "お支払期限", "Due Date"],
        "direction": "right",
        "data_type": "date",
    },
    "取引先名": {
        "labels": ["取引先名", "請求先", "宛先", "Bill To"],
        "direction": "right_or_below",
        "data_type": "text",
        "exclude_self": True,  # ラベル自体は値にしない
    },
    "件名": {
        "labels": ["件名", "Subject"],
        "direction": "right",
        "data_type": "text",
    },
    "合計金額": {
        "labels": ["合計金額", "ご請求金額", "請求金額", "合計", "Total"],
        "direction": "right",
        "data_type": "number",
    },
    "小計": {
        "labels": ["小計", "Subtotal"],
        "direction": "right",
        "data_type": "number",
    },
    "消費税": {
        "labels": ["消費税", "消費税(10%)", "Tax", "VAT"],
        "direction": "right",
        "data_type": "number",
    },
}

# CSV出力時の列順
CSV_COLUMNS = [
    "ファイル名",
    "請求書番号",
    "発行日",
    "取引先名",
    "件名",
    "小計",
    "消費税",
    "合計金額",
    "支払期限",
]

# 「取引先名」のフォールバック処理用
# 「御中」を含むセルを優先的に探す設定
HONORIFIC_KEYWORDS = ["御中", "様"]
