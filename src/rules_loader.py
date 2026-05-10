"""
YAML形式のルールファイルを読み込むモジュール
"""
import yaml
from pathlib import Path
from typing import Dict, Any


class RulesLoadError(Exception):
    """ルールファイル読込エラー"""
    pass


def load_rules(rules_path: str) -> Dict[str, Any]:
    """
    YAMLルールファイルを読み込んで辞書として返す

    Args:
        rules_path: ルールファイルのパス

    Returns:
        {
            "extraction_rules": {...},
            "csv_columns": [...],
            "fallback": {...}
        }

    Raises:
        RulesLoadError: ファイル読込・パースに失敗
    """
    path = Path(rules_path)

    if not path.exists():
        raise RulesLoadError(f"ルールファイルが見つかりません: {path}")

    if not path.is_file():
        raise RulesLoadError(f"パスがファイルではありません: {path}")

    try:
        with open(path, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f)
    except yaml.YAMLError as e:
        raise RulesLoadError(f"YAMLパースエラー ({path}): {e}")
    except Exception as e:
        raise RulesLoadError(f"ファイル読込エラー ({path}): {e}")

    if not isinstance(data, dict):
        raise RulesLoadError(f"ルールファイルの形式が不正です: {path}")

    # 必須キー検証
    if "extraction_rules" not in data:
        raise RulesLoadError(f"'extraction_rules' が定義されていません: {path}")

    if "csv_columns" not in data:
        raise RulesLoadError(f"'csv_columns' が定義されていません: {path}")

    # ファイル名は常にCSV出力に含める
    if "ファイル名" not in data["csv_columns"]:
        data["csv_columns"].insert(0, "ファイル名")

    # fallback はオプショナル(なければデフォルト値)
    data.setdefault("fallback", {})
    data["fallback"].setdefault("customer_honorific_keywords", ["御中", "様"])
    data["fallback"].setdefault("total_use_max_amount", True)
    data["fallback"].setdefault("total_calculate_from_subtotal", True)

    return data


def validate_rules(rules: Dict[str, Any]) -> list:
    """
    ルール内容の妥当性チェック
    エラーメッセージのリストを返す(空なら問題なし)
    """
    errors = []
    valid_directions = {"right", "below", "right_or_below"}
    valid_data_types = {"text", "number", "date"}

    extraction_rules = rules.get("extraction_rules", {})
    if not extraction_rules:
        errors.append("抽出ルールが1つも定義されていません")
        return errors

    for field_name, rule in extraction_rules.items():
        if not isinstance(rule, dict):
            errors.append(f"'{field_name}' のルール定義が不正です")
            continue

        labels = rule.get("labels", [])
        if not labels or not isinstance(labels, list):
            errors.append(f"'{field_name}' の labels が空または不正です")

        direction = rule.get("direction")
        if direction not in valid_directions:
            errors.append(
                f"'{field_name}' の direction が不正です: '{direction}' "
                f"(有効値: {valid_directions})"
            )

        data_type = rule.get("data_type")
        if data_type not in valid_data_types:
            errors.append(
                f"'{field_name}' の data_type が不正です: '{data_type}' "
                f"(有効値: {valid_data_types})"
            )

    # CSV列がextraction_rulesに存在するかチェック
    csv_columns = rules.get("csv_columns", [])
    for col in csv_columns:
        if col == "ファイル名":
            continue  # ファイル名は特別扱い
        if col not in extraction_rules:
            errors.append(
                f"csv_columns の '{col}' が extraction_rules に定義されていません"
            )

    return errors
