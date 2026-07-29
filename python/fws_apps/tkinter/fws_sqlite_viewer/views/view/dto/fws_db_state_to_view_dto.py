"""
Summary:
    Logicから画面へのデータベース接続状態およびテーブル一覧表示用DTOを定義するモジュール。
"""

from dataclasses import dataclass


@dataclass(frozen=True)
class FwsDbStateToViewDTO:
    """
    Summary:
        Logicから画面へのデータベース接続状態およびテーブル一覧表示用DTOクラス。
    """

    db_path: str
    """str - 現在接続されているデータベースの絶対パス。"""

    tables: list[str]
    """list[str] - データベース内に存在するテーブル名のリスト。"""
