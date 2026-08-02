"""
Summary:
    Logicから画面へのディレクトリ一覧表示用DTOを定義するモジュール。
"""

from dataclasses import dataclass


@dataclass(frozen=True)
class FwsDirInfoToViewDTO:
    """
    Summary:
        Logicから画面へのディレクトリ一覧表示用DTOクラス。
    """

    dir_path: str
    """str - 取得したディレクトリの絶対パス。"""

    items: list[tuple]
    """list[tuple] - 取得したファイル/ディレクトリ情報 (名前, 種類, サイズ, 最終更新日時) のリスト。"""

    message: str
    """str - 実行ステータスやエラー文言などのメッセージ。"""

    is_success: bool
    """bool - ディレクトリ読み込みが成功したかどうか。"""
