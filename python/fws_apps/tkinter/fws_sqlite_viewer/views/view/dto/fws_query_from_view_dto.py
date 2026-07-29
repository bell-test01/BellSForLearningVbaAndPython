"""
Summary:
    画面からLogicへのSQLクエリ実行などの入力要求用DTOを定義するモジュール。
"""

from dataclasses import dataclass


@dataclass(frozen=True)
class FwsQueryFromViewDTO:
    """
    Summary:
        画面からLogicへのSQLクエリ実行などの入力要求用DTOクラス。
    """

    db_path: str
    """str - 接続対象のSQLiteデータベースファイルの絶対パス。"""

    query: str
    """str - 実行対象のSQLクエリ文字列。"""
