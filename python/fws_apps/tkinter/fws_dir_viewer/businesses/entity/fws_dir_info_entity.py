"""
Summary:
    ディレクトリ構成情報や結果を表現するビジネスエンティティモジュール。
"""

from dataclasses import dataclass

import fws_apps.tkinter.fws_dir_viewer.views.dto.fws_dir_info_from_view_dto as fws_dir_info_from_view_dto
import fws_apps.tkinter.fws_dir_viewer.views.dto.fws_dir_info_to_view_dto as fws_dir_info_to_view_dto


@dataclass
class FwsDirInfoEntity:
    """
    Summary:
        ディレクトリ構成情報および取得結果をカプセル化するビジネスエンティティクラス。
    """

    dir_path: str = ""
    """str - ディレクトリファイルの絶対パス。"""

    items: list[tuple] = None
    """list[tuple] - 取得したファイル/ディレクトリ情報のリスト。"""

    message: str = ""
    """str - 処理時のステータスまたはエラーメッセージ。"""

    is_success: bool = True
    """bool - 処理に成功したかどうか。"""

    def __post_init__(self):
        if self.items is None:
            self.items = []

    @classmethod
    def from_view_dto(cls, dto: fws_dir_info_from_view_dto.FwsDirInfoFromViewDTO) -> "FwsDirInfoEntity":
        """
        Summary:
            入力用DTOからディレクトリ情報エンティティを生成します。
        Args:
            dto: FwsDirInfoFromViewDTO - 画面から送信された入力DTO。
        Returns:
            FwsDirInfoEntity - 生成されたエンティティ。
        """
        return cls(dir_path=dto.dir_path)

    def to_view_dto(self) -> fws_dir_info_to_view_dto.FwsDirInfoToViewDTO:
        """
        Summary:
            ディレクトリ情報結果を表示用DTOに変換します。
        Returns:
            FwsDirInfoToViewDTO - 画面表示用DTO。
        """
        return fws_dir_info_to_view_dto.FwsDirInfoToViewDTO(dir_path=self.dir_path, items=self.items, message=self.message, is_success=self.is_success)
