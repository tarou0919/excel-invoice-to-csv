"""
Excel請求書一括CSV化ツール
使い方:
    python src/main.py                    # samples/ → output/invoices.csv
    python src/main.py 入力Dir 出力CSV    # カスタムパス
"""
import sys
from pathlib import Path

# srcディレクトリをパスに追加
sys.path.insert(0, str(Path(__file__).parent))

from extractor import extract_from_file
from csv_writer import write_csv


def collect_excel_files(input_dir):
    """指定ディレクトリから .xlsx ファイルを収集"""
    input_dir = Path(input_dir)
    if not input_dir.exists():
        print(f"❌ 入力ディレクトリが見つかりません: {input_dir}")
        return []
    files = sorted(input_dir.glob("*.xlsx"))
    # ~$から始まる一時ファイルは除外(Excel開いてる時にできるやつ)
    files = [f for f in files if not f.name.startswith("~$")]
    return files


def process(input_dir, output_csv):
    """メイン処理"""
    print(f"📂 入力フォルダ: {input_dir}")
    print(f"📄 出力CSV: {output_csv}")
    print()

    files = collect_excel_files(input_dir)
    if not files:
        print("⚠️  処理対象のExcelファイルが見つかりませんでした。")
        return

    print(f"🔍 {len(files)} 件のExcelファイルを処理します...\n")

    records = []
    success = 0
    for i, file_path in enumerate(files, 1):
        print(f"[{i}/{len(files)}] {file_path.name}")
        record = extract_from_file(file_path)
        records.append(record)

        # 抽出結果のサマリー表示
        invoice_no = record.get("請求書番号") or "(未取得)"
        total = record.get("合計金額")
        total_str = f"¥{int(total):,}" if total else "(未取得)"
        customer = record.get("取引先名") or "(未取得)"
        print(f"        請求書番号: {invoice_no}")
        print(f"        取引先: {customer}")
        print(f"        合計: {total_str}")

        # 必須項目が取れていれば success カウント
        if record.get("請求書番号") and record.get("合計金額"):
            success += 1

    # CSV書き出し
    output_path = write_csv(records, output_csv)

    print()
    print("=" * 50)
    print(f"✅ 完了: {output_path}")
    print(f"   処理件数: {len(records)} 件")
    print(f"   主要項目取得成功: {success} / {len(records)} 件")
    print("=" * 50)


def main():
    # 引数処理
    if len(sys.argv) >= 3:
        input_dir = sys.argv[1]
        output_csv = sys.argv[2]
    else:
        # デフォルト
        base = Path(__file__).parent.parent
        input_dir = base / "samples"
        output_csv = base / "output" / "invoices.csv"

    process(input_dir, output_csv)


if __name__ == "__main__":
    main()
