import unittest
from unittest.mock import MagicMock

import fws_apps.tkinter.fws_spec_generator.models.dto.fws_generator_from_view_dto as fws_generator_from_view_dto
import fws_apps.tkinter.fws_spec_generator.models.dto.fws_generator_to_view_dto as fws_generator_to_view_dto
import fws_apps.tkinter.fws_spec_generator.views.event.fws_app_events as fws_app_events
from fws_apps.tkinter.fws_spec_generator.models.entity.fws_generator_entity import FwsGeneratorEntity


class TestFwsApp(unittest.TestCase):
    """MagicMock を使用した単体テスト"""

    def setUp(self):
        self.mock_view = MagicMock()
        self.mock_logic = MagicMock()

        self.app_events = fws_app_events.FwsAppEvents(
            test_view=self.mock_view, test_logic=self.mock_logic
        )

    def test_handle_generate_success(self):
        """生成処理の正常系テスト"""
        input_dto = fws_generator_from_view_dto.FwsGeneratorFromViewDTO(
            source_dir="c:/some/source", output_dir="c:/some/output"
        )
        self.mock_view.get_input_dto.return_value = input_dto
        self.mock_logic.execute_generation.return_value = "c:/some/output/index.html"

        # ダミーのモック
        import os
        orig_exists = os.path.exists
        os.path.exists = MagicMock(return_value=True)

        try:
            self.app_events.handle_generate()
            self.mock_logic.update_paths.assert_called_once_with(input_dto)
            self.mock_logic.execute_generation.assert_called_once()
            self.mock_view.show_info.assert_called_once()
        finally:
            os.path.exists = orig_exists

    def test_refresh_view(self):
        """描画更新時の DTO 横流しテスト"""
        display_dto = fws_generator_to_view_dto.FwsGeneratorToViewDTO(
            source_dir="c:/some/source",
            output_dir="c:/some/output",
            log_text="ログです",
            status_message="ステータスです"
        )
        self.mock_logic.get_to_view_dto.return_value = display_dto

        self.app_events._refresh_view()

        self.mock_view.update_view.assert_called_once_with(display_dto)

    def test_generator_entity_dto_conversion(self):
        """Entity と DTO の相互変換テスト"""
        entity = FwsGeneratorEntity(
            id="123",
            source_dir="dir_a",
            output_dir="dir_b",
            log_text="text_c",
            status_message="msg_d"
        )
        dto = entity.to_view_dto()
        self.assertEqual(dto.source_dir, "dir_a")
        self.assertEqual(dto.output_dir, "dir_b")
        self.assertEqual(dto.log_text, "text_c")
        self.assertEqual(dto.status_message, "msg_d")


if __name__ == "__main__":
    unittest.main()
