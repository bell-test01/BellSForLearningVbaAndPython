"""
Summary:
    SQLite Viewer アプリケーションの Logic 層に関するモック単体テストモジュール。
"""

import os
import sys
import unittest
from unittest.mock import MagicMock

# プロジェクトルートのパス解決
project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), "../../../../"))
if project_root not in sys.path:
    sys.path.append(project_root)

import fws_apps.tkinter.fws_sqlite_viewer.logics.fws_app_logic as fws_app_logic
import fws_apps.tkinter.fws_sqlite_viewer.models.dto.fws_query_from_view_dto as fws_query_from_view_dto


class FwsTestApp(unittest.TestCase):
    """
    Summary:
        FwsAppLogic クラスの各メソッドに対し、モック接続オブジェクトを用いた単体テストを行うクラス。
    """

    def setUp(self):
        # 接続、カーソルのモックをセットアップ
        self.mock_conn = MagicMock()
        self.mock_cursor = MagicMock()
        self.mock_conn.cursor.return_value = self.mock_cursor

        # テスト対象Logicインスタンスを、モックを注入して生成
        self.logic = fws_app_logic.FwsAppLogic(test_sqlite_conn=self.mock_conn)

    def test_create_new_db_success(self):
        """
        Summary:
            新規にデータベースファイルを作成・接続し、初期状態（テーブルなし）を正しく取得できることをテストします。
        """
        # 初回はテーブルなし
        self.mock_cursor.fetchall.return_value = []

        original_exists = os.path.exists
        os.path.exists = lambda path: True

        try:
            state_dto = self.logic.create_new_db("new_dummy_db.sqlite")

            self.assertEqual(state_dto.db_path, "new_dummy_db.sqlite")
            self.assertEqual(state_dto.tables, [])
            self.mock_conn.cursor.assert_called()
        finally:
            os.path.exists = original_exists

    def test_get_db_state_success(self):
        """
        Summary:
            データベース内のテーブル名一覧が正しく取得できることをテストします。
        """
        # モックの振る舞い設定 (users と items テーブルが存在する想定)
        self.mock_cursor.fetchall.return_value = [("users",), ("items",)]

        # DBファイルが存在する前提のパス（テスト用なので実際には開かないが、os.path.existsをダミーで通すか、実在するファイルを指定）
        # os.path.exists の結果をモックするか、あるいはプロジェクト内の実在ファイルを使用する。
        # ここでは setUp またはテスト内で os.path.exists をモック化するのが安全。
        original_exists = os.path.exists
        os.path.exists = lambda path: True

        try:
            state_dto = self.logic.get_db_state("dummy_db.sqlite")

            self.assertEqual(state_dto.db_path, "dummy_db.sqlite")
            self.assertEqual(state_dto.tables, ["users", "items"])
            self.mock_cursor.execute.assert_called_once_with(
                "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%';"
            )
        finally:
            os.path.exists = original_exists

    def test_execute_query_select_success(self):
        """
        Summary:
            SELECTクエリが成功し、列名と行データが正しく取得されることをテストします。
        """
        # SELECTクエリ用のモック記述
        self.mock_cursor.description = (("id", None, None, None, None, None, None), ("name", None, None, None, None, None, None))
        self.mock_cursor.fetchall.return_value = [(1, "Alice"), (2, "Bob")]

        original_exists = os.path.exists
        os.path.exists = lambda path: True

        try:
            dto = fws_query_from_view_dto.FwsQueryFromViewDTO(
                db_path="dummy_db.sqlite",
                query="SELECT id, name FROM users;"
            )
            result_dto = self.logic.execute_query(dto)

            self.assertTrue(result_dto.is_success)
            self.assertEqual(result_dto.columns, ["id", "name"])
            self.assertEqual(result_dto.rows, [(1, "Alice"), (2, "Bob")])
            self.assertIn("2 件取得しました", result_dto.message)
            self.mock_cursor.execute.assert_called_once_with("SELECT id, name FROM users;")
        finally:
            os.path.exists = original_exists

    def test_execute_query_update_success(self):
        """
        Summary:
            更新系クエリ（INSERT/UPDATE等）が成功し、commitが呼び出されることをテストします。
        """
        # 更新系クエリ（descriptionがNoneになる）のモック設定
        self.mock_cursor.description = None
        self.mock_cursor.rowcount = 1

        original_exists = os.path.exists
        os.path.exists = lambda path: True

        try:
            dto = fws_query_from_view_dto.FwsQueryFromViewDTO(
                db_path="dummy_db.sqlite",
                query="INSERT INTO users (name) VALUES ('Charlie');"
            )
            result_dto = self.logic.execute_query(dto)

            self.assertTrue(result_dto.is_success)
            self.assertEqual(result_dto.columns, [])
            self.assertEqual(result_dto.rows, [])
            self.assertIn("影響した行数: 1", result_dto.message)
            self.mock_conn.commit.assert_called_once()
        finally:
            os.path.exists = original_exists

    def test_execute_query_failure(self):
        """
        Summary:
            クエリ実行時にエラーが発生した場合に、正しくエラーメッセージを取得できることをテストします。
        """
        # 例外を投げるように設定
        self.mock_cursor.execute.side_effect = Exception("Syntax Error")

        original_exists = os.path.exists
        os.path.exists = lambda path: True

        try:
            dto = fws_query_from_view_dto.FwsQueryFromViewDTO(
                db_path="dummy_db.sqlite",
                query="INVALID SQL;"
            )
            result_dto = self.logic.execute_query(dto)

            self.assertFalse(result_dto.is_success)
            self.assertIn("Syntax Error", result_dto.message)
        finally:
            os.path.exists = original_exists


if __name__ == "__main__":
    unittest.main()
