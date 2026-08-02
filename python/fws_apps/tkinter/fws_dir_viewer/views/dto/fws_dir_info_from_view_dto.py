"""
Summary:
    画面からLogicへのディレクトリパス入力要求用DTOを定義するモジュール。
"""

from dataclasses import dataclass


@dataclass(frozen=True)
class FwsDirInfoFromViewDTO:
    """
    Summary:
        画面からLogicへのディレクトリパス入力要求用DTOクラス。
    """

    dir_path: str
    """str - 読み込み対象のディレクトリの絶対パス。"""
