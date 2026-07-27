"""
Summary:
    データベースクエリの実行情報や結果を表すビジネスエンティティモジュール。
"""

from dataclasses import dataclass

import fws_apps.tkinter.fws_sqlite_viewer.models.dto.fws_db_state_to_view_dto as fws_db_state_to_view_dto
import fws_apps.tkinter.fws_sqlite_viewer.models.dto.fws_query_from_view_dto as fws_query_from_view_dto
import fws_apps.tkinter.fws_sqlite_viewer.models.dto.fws_query_result_to_view_dto as fws_query_result_to_view_dto


@dataclass
class FwsQueryEntity:
    """
    Summary:
        データベースクエリ情報および実行結果をカプセル化するビジネスエンティティクラス。
    """

    db_path: str
    """str - データベースファイルの絶対パス。"""

    query: str = ""
    """str - 実行対象のSQLクエリ文字列。"""

    columns: list[str] = None
    """list[str] - クエリ実行で取得したSELECTの列名リスト。"""

    rows: list[tuple] = None
    """list[tuple] - クエリ実行で取得したレコードリスト。"""

    message: str = ""
    """str - クエリ実行時のステータスまたはエラーメッセージ。"""

    is_success: bool = True
    """bool - クエリ実行に成功したかどうか。"""

    tables: list[str] = None
    """list[str] - データベース内のテーブル名リスト。"""

    def __post_init__(self):
        if self.columns is None:
            self.columns = []
        if self.rows is None:
            self.rows = []
        if self.tables is None:
            self.tables = []

    @classmethod
    def from_view_dto(
        cls, dto: fws_query_from_view_dto.FwsQueryFromViewDTO
    ) -> "FwsQueryEntity":
        """
        Summary:
            入力用DTOからクエリエンティティを生成します。
        Args:
            dto: FwsQueryFromViewDTO - 画面から送信された入力DTO。
        Returns:
            FwsQueryEntity - 生成されたエンティティ。
        """
        return cls(
            db_path=dto.db_path,
            query=dto.query,
        )

    def to_result_view_dto(
        self,
    ) -> fws_query_result_to_view_dto.FwsQueryResultToViewDTO:
        """
        Summary:
            クエリ実行結果を表示用DTOに変換します。
        Returns:
            FwsQueryResultToViewDTO - 画面表示用DTO。
        """
        return fws_query_result_to_view_dto.FwsQueryResultToViewDTO(
            columns=self.columns,
            rows=self.rows,
            message=self.message,
            is_success=self.is_success,
        )

    def to_state_view_dto(
        self,
    ) -> fws_db_state_to_view_dto.FwsDbStateToViewDTO:
        """
        Summary:
            データベーススキーマの接続状態を表示用DTOに変換します。
        Returns:
            FwsDbStateToViewDTO - データベース状態DTO。
        """
        return fws_db_state_to_view_dto.FwsDbStateToViewDTO(
            db_path=self.db_path,
            tables=self.tables,
        )
