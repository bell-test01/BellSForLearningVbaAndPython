"""
Summary:
    画面からの要求を受け取り、ビジネスロジックを呼び出して結果をDTOで返すロジックモジュール。
"""

import fws_apps.tkinter.fws_dir_viewer.views.dto.fws_dir_info_from_view_dto as fws_dir_info_from_view_dto
import fws_apps.tkinter.fws_dir_viewer.views.dto.fws_dir_info_to_view_dto as fws_dir_info_to_view_dto
import fws_apps.tkinter.fws_dir_viewer.businesses.entity.fws_dir_info_entity as fws_dir_info_entity
import fws_apps.tkinter.fws_dir_viewer.businesses.business.fws_dir_info_business as fws_dir_info_business


class FwsAppLogic:
    """
    Summary:
        ディレクトリ構成取得のフローを制御するロジッククラス。
    """

    def __init__(self, test_business=None):
        self.business = test_business if test_business is not None else fws_dir_info_business.FwsDirInfoBusiness()
        """FwsDirInfoBusiness - ディレクトリ情報取得を行うビジネスクラス。"""

    def load_directory(self, dto: fws_dir_info_from_view_dto.FwsDirInfoFromViewDTO) -> fws_dir_info_to_view_dto.FwsDirInfoToViewDTO:
        """
        Summary:
            画面から受け取ったパス入力DTOを元に、Businessクラスでディレクトリ情報を取得し、結果を表示用DTOとして返します。
        Args:
            dto: FwsDirInfoFromViewDTO - 実行要求パス情報。
        Returns:
            FwsDirInfoToViewDTO - 実行結果DTO。
        """
        entity = fws_dir_info_entity.FwsDirInfoEntity.from_view_dto(dto)
        result_entity = self.business.get_dir_info(entity)
        return result_entity.to_view_dto()
