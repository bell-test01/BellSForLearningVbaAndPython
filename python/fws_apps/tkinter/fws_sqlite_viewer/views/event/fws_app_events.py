"""
Summary:
    SQLite Viewer画面におけるユーザーアクション制御およびビジネスロジック連携を担当するモジュール。
"""

from tkinter import filedialog

import fws_apps.tkinter.fws_sqlite_viewer.logics.fws_app_logic as fws_app_logic
import fws_apps.tkinter.fws_sqlite_viewer.views.view.fws_app_view as fws_app_view


class FwsAppEvents:
    """
    Summary:
        ビューからのイベントおよびビジネスロジックのフロー制御を担当するクラス。
    """

    def __init__(self, test_view=None, test_logic=None):
        self.view = test_view if test_view is not None else fws_app_view.FwsAppView()
        """FwsAppView - 表示制御用ビューインスタンス。"""

        self.logic = test_logic if test_logic is not None else fws_app_logic.FwsAppLogic()
        """FwsAppLogic - ビジネスロジックインスタンス。"""

        self._setup_event_bindings()

    def _setup_event_bindings(self):
        """
        Summary:
            各種ボタンやキーボード操作に対して、Direct Event Binding でイベントハンドラーをバインドします。
        """
        self.view.select_db_button.config(command=self.handle_select_db)
        self.view.execute_button.config(command=self.handle_execute_query)
        self.view.table_tree.bind("<Double-1>", self.handle_table_double_click)

        # ファイルメニューへのバインド
        self.view.file_menu.entryconfigure("新規データベース作成...", command=self.handle_create_db)
        self.view.file_menu.entryconfigure("データベースを開く...", command=self.handle_select_db)
        self.view.file_menu.entryconfigure("再読み込み", command=self.handle_reload_db)

        # ボタンのバインド
        self.view.reload_db_button.config(command=self.handle_reload_db)

        # エディタのショートカットキーバインド (Ctrl+A 全選択)
        self.view.sql_text.text_widget.bind("<Control-a>", self.handle_select_all)
        self.view.sql_text.text_widget.bind("<Control-A>", self.handle_select_all)

    # ==========================================
    # イベントハンドラー群
    # ==========================================

    def handle_select_all(self, event):
        """
        Summary:
            SQL入力エリアで Ctrl+A が押された際、全テキストを選択します。
        Args:
            event: tkinter.Event - イベントオブジェクト。
        Returns:
            str - 'break' を返して標準挙動を抑制し多重選択を防ぎます。
        """
        self.view.sql_text.text_widget.tag_add("sel", "1.0", "end")
        return "break"

    def handle_reload_db(self):
        """
        Summary:
            接続中のデータベースを再読み込みし、テーブル一覧を更新します。
        UserAction:
            再読み込みボタンまたはメニューをクリックした時 - 現在接続されているデータベースからテーブル一覧を再取得して表示を更新する。
        """
        input_dto = self.view.get_input_dto()
        file_path = input_dto.db_path
        if not file_path:
            self.view.set_status_message("エラー: 再読み込みするデータベースが選択されていません。")
            return

        db_state = self.logic.get_db_state(file_path)
        self.view.set_db_state(db_state)
        self.view.set_status_message(f"データベースを再読み込みしました: {file_path}")

    def handle_create_db(self):
        """
        Summary:
            新規データベース作成メニューを選択した際、保存先ダイアログを表示して新規データベースを作成・接続します。
        UserAction:
            新規データベース作成... をクリックした時 - 保存先を指定して空のデータベースファイルを作成し、接続する。
        """
        file_path = filedialog.asksaveasfilename(
            title="新規SQLiteデータベースの保存先を指定してください",
            defaultextension=".db",
            filetypes=[
                ("SQLite Databases", "*.db;*.sqlite;*.sqlite3;*.db3"),
                ("All Files", "*.*"),
            ],
        )
        if not file_path:
            return

        db_state = self.logic.create_new_db(file_path)
        self.view.set_db_state(db_state)
        self.view.set_status_message(f"新規データベースを作成・接続しました: {file_path}")

    def handle_select_db(self):
        """
        Summary:
            データベースファイル選択ボタンをクリックした際、ファイルダイアログを表示してデータベースに接続します。
        UserAction:
            ファイル参照ボタンをクリックした時 - ファイルダイアログからデータベースファイルを選択し、テーブル一覧を更新する。
        """
        file_path = filedialog.askopenfilename(
            title="SQLiteデータベースファイルを選択してください",
            filetypes=[
                ("SQLite Databases", "*.db;*.sqlite;*.sqlite3;*.db3"),
                ("All Files", "*.*"),
            ],
        )
        if not file_path:
            return

        # DB状態の取得と反映
        db_state = self.logic.get_db_state(file_path)
        self.view.set_db_state(db_state)
        self.view.set_status_message(f"データベースに接続しました: {file_path}")

    def handle_execute_query(self):
        """
        Summary:
            SQLクエリ実行ボタンがクリックされた際、入力されたSQLを実行し結果を画面に表示します。
        UserAction:
            SQLクエリ実行ボタンをクリックした時 - 入力されたSQLを実行し、実行結果をグリッドに表示する。
        """
        input_dto = self.view.get_input_dto()
        result_dto = self.logic.execute_query(input_dto)
        self.view.set_query_result(result_dto)

    def handle_table_double_click(self, event):
        """
        Summary:
            テーブル一覧内のテーブル名がダブルクリックされた際、自動でSELECTクエリを作成し実行します。
        Args:
            event: tkinter.Event - イベントオブジェクト。
        UserAction:
            テーブル一覧のテーブル名をダブルクリックした時 - SELECT * FROM [テーブル名] のクエリを自動生成して実行する。
        """
        table_name = self.view.get_selected_table()
        if not table_name:
            return

        # クエリを自動生成して入力エリアにセット
        sql_query = f"SELECT * FROM {table_name};"
        self.view.set_sql_text(sql_query)

        # 自動実行
        self.handle_execute_query()
