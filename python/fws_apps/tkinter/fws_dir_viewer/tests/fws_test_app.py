"""
Summary:
    fws_dir_viewer の単体テストモジュール。GUIやOSのI/Oに依存しないロジック層・ビジネス層をテストします。
"""

import unittest
from unittest.mock import MagicMock

import fws_apps.tkinter.fws_dir_viewer.views.dto.fws_dir_info_from_view_dto as fws_dir_info_from_view_dto
import fws_apps.tkinter.fws_dir_viewer.views.logic.fws_app_logic as fws_app_logic
import fws_apps.tkinter.fws_dir_viewer.businesses.entity.fws_dir_info_entity as fws_dir_info_entity


class TestFwsAppLogic(unittest.TestCase):
    """
    Summary:
        LogicクラスおよびBusiness層連携の単体テストクラス。
    """

    def setUp(self):
        self.mock_business = MagicMock()
        self.logic = fws_app_logic.FwsAppLogic(test_business=self.mock_business)

    def test_load_directory_success(self):
        """
        Summary:
            load_directory メソッドが正常にディレクトリ情報を取得してDTOを返すことを検証します。
        """
        input_dto = fws_dir_info_from_view_dto.FwsDirInfoFromViewDTO(dir_path="c:/dummy_path")

        mock_result_entity = fws_dir_info_entity.FwsDirInfoEntity(
            dir_path="c:/dummy_path",
            items=[("dummy.txt", "File", "1024", "2026-08-01 10:00:00", "c:/dummy_path/dummy.txt")],
            message="成功: 1 件のアイテムを取得しました。",
            is_success=True,
        )
        self.mock_business.get_dir_info.return_value = mock_result_entity

        result_dto = self.logic.load_directory(input_dto)

        self.mock_business.get_dir_info.assert_called_once()
        self.assertTrue(result_dto.is_success)
        self.assertEqual(len(result_dto.items), 1)
        self.assertEqual(result_dto.items[0][0], "dummy.txt")

    def test_load_directory_failure(self):
        """
        Summary:
            load_directory メソッドがエラー時の結果を正しく処理してDTOを返すことを検証します。
        """
        input_dto = fws_dir_info_from_view_dto.FwsDirInfoFromViewDTO(dir_path="")

        mock_result_entity = fws_dir_info_entity.FwsDirInfoEntity(
            dir_path="",
            items=[],
            message="エラー: ディレクトリパスを入力してください。",
            is_success=False,
        )
        self.mock_business.get_dir_info.return_value = mock_result_entity

        result_dto = self.logic.load_directory(input_dto)

        self.mock_business.get_dir_info.assert_called_once()
        self.assertFalse(result_dto.is_success)
        self.assertEqual(len(result_dto.items), 0)
        self.assertEqual(result_dto.message, "エラー: ディレクトリパスを入力してください。")


if __name__ == "__main__":
    unittest.main()
