"""
Excel請求書一括CSV化ツール v2.0

使い方:
    # デフォルト
    python src/main.py

    # 入力・出力パス指定
    python src/main.py samples/ output/result.csv

    # カスタムルール使用
    python src/main.py --rules rules/customer_a.yaml

    # 全部指定
    python src/main.py samples/ output/result.csv --rules rules/customer_a.yaml

    # 詳細ログ出力(デバッグ用)
    python src/main.py --verbose
"""
import sys
import logging
import argparse
from pathlib import Path
from datetime import datetime

# srcディレクトリをパスに追加
sys.path.insert(0, str(Path(__file__).parent))

from extractor import extract_from_file
from csv_writer import write_csv
from rules_loader import load_rules, validate_rules, RulesLoadError


def setup_logging(verbose: bool = False, log_file: Path = None):
    """ロギング設定"""
    level = logging.DEBUG if verbose else logging.INFO

    handlers = [logging.StreamHandler(sys.stderr)]

    if log_file:
        log_file.parent.mkdir(parents=True, exist_ok=True)
        handlers.append(logging.FileHandler(log_file, encoding="utf-8"))

    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=handlers,
        force=True,
    )

    # 通常モードではDEBUGログを抑制(extractorからのデバッグログをカット)
    if not verbose:
        logging.getLogger("extractor").setLevel(logging.WARNING)


def collect_excel_files(input_dir: Path) -> list:
    """指定ディレクトリから .xlsx ファイルを収集"""
    if not input_dir.exists():
        raise FileNotFoundError(f"入力ディレクトリが見つかりません: {input_dir}")
    if not input_dir.is_dir():
        raise NotADirectoryError(f"入力パスがディレクトリではありません: {input_dir}")

    files = sorted(input_dir.glob("*.xlsx")) + sorted(input_dir.glob("*.xlsm"))
    # ~$ から始まる一時ファイルは除外
    files = [f for f in files if not f.name.startswith("~$")]
    return files


def parse_args():
    """コマンドライン引数の解析"""
    base = Path(__file__).parent.parent

    parser = argparse.ArgumentParser(
        description="Excel請求書を一括でCSVに変換します",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用例:
    python src/main.py
    python src/main.py samples/ output/invoices.csv
    python src/main.py --rules rules/customer_a.yaml
    python src/main.py samples/ output.csv --rules rules/custom.yaml --verbose
        """,
    )
    parser.add_argument(
        "input_dir",
        nargs="?",
        default=str(base / "samples"),
        help="入力Excelファイルのディレクトリ (デフォルト: samples/)",
    )
    parser.add_argument(
        "output_csv",
        nargs="?",
        default=str(base / "output" / "invoices.csv"),
        help="出力CSVファイルのパス (デフォルト: output/invoices.csv)",
    )
    parser.add_argument(
        "--rules",
        "-r",
        default=str(base / "rules" / "default.yaml"),
        help="抽出ルールYAMLファイル (デフォルト: rules/default.yaml)",
    )
    parser.add_argument(
        "--verbose",
        "-v",
        action="store_true",
        help="詳細ログを表示(デバッグ用)",
    )
    parser.add_argument(
        "--log-file",
        default=None,
        help="ログファイルの出力先(指定なしならファイル出力なし)",
    )
    return parser.parse_args()


def process(input_dir: Path, output_csv: Path, rules: dict, verbose: bool = False):
    """メイン処理"""
    print(f"📂 入力フォルダ: {input_dir}")
    print(f"📄 出力CSV: {output_csv}")
    print(f"📋 ルールファイル: {rules.get('_path', '(指定なし)')}")
    print()

    csv_columns = rules["csv_columns"]

    # ファイル収集
    try:
        files = collect_excel_files(input_dir)
    except (FileNotFoundError, NotADirectoryError) as e:
        print(f"❌ エラー: {e}")
        return 1

    if not files:
        print("⚠️  処理対象のExcelファイル(.xlsx/.xlsm)が見つかりませんでした。")
        return 1

    print(f"🔍 {len(files)} 件のExcelファイルを処理します...\n")

    records = []
    success = 0
    failed = 0

    for i, file_path in enumerate(files, 1):
        print(f"[{i}/{len(files)}] {file_path.name}")
        try:
            record = extract_from_file(file_path, rules)
        except Exception as e:
            print(f"        ❌ 処理失敗: {type(e).__name__}: {e}")
            failed += 1
            records.append({"ファイル名": file_path.name})
            continue

        # エラーがあった場合
        if record.get("_error"):
            print(f"        ⚠️  {record['_error']}")
            failed += 1

        # 結果サマリー表示(主要3項目)
        invoice_no = record.get("請求書番号") or "(未取得)"
        total = record.get("合計金額")
        total_str = f"¥{int(total):,}" if total else "(未取得)"
        customer = record.get("取引先名") or "(未取得)"
        print(f"        請求書番号: {invoice_no}")
        print(f"        取引先: {customer}")
        print(f"        合計: {total_str}")

        # _error フィールドはCSVに出力しないので削除
        record.pop("_error", None)
        records.append(record)

        # 主要項目が取得できていれば成功
        if record.get("請求書番号") and record.get("合計金額"):
            success += 1

    # CSV書き出し
    try:
        output_path = write_csv(records, output_csv, csv_columns)
    except Exception as e:
        print(f"\n❌ CSV書き出し失敗: {e}")
        return 1

    print()
    print("=" * 60)
    print(f"✅ 完了: {output_path}")
    print(f"   処理件数: {len(records)} 件")
    print(f"   主要項目取得成功: {success} / {len(records)} 件")
    if failed > 0:
        print(f"   ⚠️  エラー発生: {failed} 件")
    print("=" * 60)

    return 0


def main():
    args = parse_args()

    log_file = Path(args.log_file) if args.log_file else None
    setup_logging(args.verbose, log_file)

    # ルールファイル読込
    try:
        rules = load_rules(args.rules)
        rules["_path"] = args.rules
    except RulesLoadError as e:
        print(f"❌ ルール読込エラー: {e}")
        return 1

    # ルール検証
    errors = validate_rules(rules)
    if errors:
        print("❌ ルールファイルに問題があります:")
        for err in errors:
            print(f"   - {err}")
        return 1

    # 処理実行
    return process(
        Path(args.input_dir),
        Path(args.output_csv),
        rules,
        args.verbose,
    )


if __name__ == "__main__":
    sys.exit(main())
