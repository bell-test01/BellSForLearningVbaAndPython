"""
Summary:
    SQLiteデータベースの操作、SQLクエリ実行、スキーマ情報の取得などを行うビジネスロジックモジュール。
"""

import os
import sqlite3

import fws_apps.tkinter.fws_sqlite_viewer.models.dto.fws_db_state_to_view_dto as fws_db_state_to_view_dto
import fws_apps.tkinter.fws_sqlite_viewer.models.dto.fws_query_from_view_dto as fws_query_from_view_dto
import fws_apps.tkinter.fws_sqlite_viewer.models.dto.fws_query_result_to_view_dto as fws_query_result_to_view_dto
import fws_apps.tkinter.fws_sqlite_viewer.models.entity.fws_query_entity as fws_query_entity


class FwsAppLogic:
    """
    Summary:
        SQLiteデータベースへの接続、SQL実行、テーブル情報ロードなどのビジネスロジックを制御するクラス。
    """

    def __init__(self, test_sqlite_conn=None):
        self.test_sqlite_conn = test_sqlite_conn
        """sqlite3.Connection | None - テスト用の接続モックオブジェクト。"""

    def create_new_db(self, db_path: str) -> fws_db_state_to_view_dto.FwsDbStateToViewDTO:
        """
        Summary:
            新規にSQLiteデータベースファイルを作成し、状態表示用DTOを返します。
        Args:
            db_path: str - 新規作成するデータベースファイルの絶対パス。
        Returns:
            FwsDbStateToViewDTO - データベース接続状態DTO。
        """
        if not db_path:
            entity = fws_query_entity.FwsQueryEntity(
                db_path="",
                tables=[],
            )
            return entity.to_state_view_dto()

        conn = None
        try:
            conn = self._get_connection(db_path)
        finally:
            if conn is not None and self.test_sqlite_conn is None:
                conn.close()

        return self.get_db_state(db_path)

    def get_db_state(self, db_path: str) -> fws_db_state_to_view_dto.FwsDbStateToViewDTO:
        """
        Summary:
            接続先データベースのテーブル名一覧を取得し、状態表示用DTOを返します。
        Args:
            db_path: str - データベースの絶対パス。
        Returns:
            FwsDbStateToViewDTO - データベース接続状態DTO。
        """
        if not db_path or not os.path.exists(db_path):
            entity = fws_query_entity.FwsQueryEntity(
                db_path=db_path,
                tables=[],
            )
            return entity.to_state_view_dto()

        conn = None
        try:
            conn = self._get_connection(db_path)
            cursor = conn.cursor()
            cursor.execute(
                "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%';"
            )
            tables = [row[0] for row in cursor.fetchall()]
            entity = fws_query_entity.FwsQueryEntity(
                db_path=db_path,
                tables=tables,
            )
            return entity.to_state_view_dto()
        except Exception as e:
            entity = fws_query_entity.FwsQueryEntity(
                db_path=db_path,
                tables=[],
            )
            return entity.to_state_view_dto()
        finally:
            if conn is not None and self.test_sqlite_conn is None:
                conn.close()

    def execute_query(
        self, dto: fws_query_from_view_dto.FwsQueryFromViewDTO
    ) -> fws_query_result_to_view_dto.FwsQueryResultToViewDTO:
        """
        Summary:
            画面から受け取ったクエリ実行DTOに基づき、SQLを実行して結果を表示用DTOとして返します。
        Args:
            dto: FwsQueryFromViewDTO - 実行要求クエリの情報。
        Returns:
            FwsQueryResultToViewDTO - クエリの実行結果DTO。
        """
        db_path = dto.db_path
        query = dto.query.strip()

        if not db_path or not os.path.exists(db_path):
            entity = fws_query_entity.FwsQueryEntity(
                db_path=db_path,
                query=query,
                is_success=False,
                message="エラー: 有効なデータベースファイルを選択してください。",
            )
            return entity.to_result_view_dto()

        if not query:
            entity = fws_query_entity.FwsQueryEntity(
                db_path=db_path,
                query=query,
                is_success=False,
                message="エラー: SQLを入力してください。",
            )
            return entity.to_result_view_dto()

        conn = None
        try:
            conn = self._get_connection(db_path)
            cursor = conn.cursor()
            cursor.execute(query)

            # SELECT クエリか、あるいは更新系クエリかの判定
            description = cursor.description
            if description is not None:
                # 検索クエリの場合
                columns = [desc[0] for desc in description]
                rows = cursor.fetchall()
                entity = fws_query_entity.FwsQueryEntity(
                    db_path=db_path,
                    query=query,
                    columns=columns,
                    rows=rows,
                    is_success=True,
                    message=f"成功: {len(rows)} 件取得しました。",
                )
            else:
                # 更新クエリ（INSERT/UPDATE/DELETE/CREATEなど）の場合
                conn.commit()
                rowcount = cursor.rowcount
                entity = fws_query_entity.FwsQueryEntity(
                    db_path=db_path,
                    query=query,
                    columns=[],
                    rows=[],
                    is_success=True,
                    message=f"成功: クエリを実行しました（影響した行数: {rowcount}）。",
                )
            return entity.to_result_view_dto()

        except Exception as e:
            if conn is not None and self.test_sqlite_conn is None:
                conn.rollback()
            entity = fws_query_entity.FwsQueryEntity(
                db_path=db_path,
                query=query,
                is_success=False,
                message=f"実行エラー: {str(e)}",
            )
            return entity.to_result_view_dto()
        finally:
            if conn is not None and self.test_sqlite_conn is None:
                conn.close()

    def _get_connection(self, db_path: str) -> sqlite3.Connection:
        """
        Summary:
            SQLiteコネクションを取得します。テスト用接続オブジェクトがある場合はそれを返します。
        Args:
            db_path: str - データベースパス。
        Returns:
            sqlite3.Connection - データベース接続オブジェクト。
        """
        if self.test_sqlite_conn is not None:
            return self.test_sqlite_conn
        return sqlite3.connect(db_path)
