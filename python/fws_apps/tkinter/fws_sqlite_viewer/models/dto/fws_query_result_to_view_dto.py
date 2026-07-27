"""
Summary:
    Logicから画面へのクエリ実行結果表示用DTOを定義するモジュール。
"""

from dataclasses import dataclass


@dataclass(frozen=True)
class FwsQueryResultToViewDTO:
    """
    Summary:
        Logicから画面へのクエリ実行結果表示用DTOクラス。
    """

    columns: list[str]
    """list[str] - SELECT結果の列名リスト。"""

    rows: list[tuple]
    """list[tuple] - SELECT結果の行データリスト。"""

    message: str
    """str - 実行ステータスやエラー文言などのメッセージ。"""

    is_success: bool
    """bool - クエリの実行が成功したかどうか。"""
