"""
Summary:
    ディレクトリビューアー画面におけるユーザーアクション制御およびビジネスロジック連携を担当するモジュール。
"""

from tkinter import filedialog

import fws_apps.tkinter.fws_dir_viewer.views.logic.fws_app_logic as fws_app_logic
import fws_apps.tkinter.fws_dir_viewer.views.view.fws_app_view as fws_app_view


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
        self.view.select_dir_button.config(command=self.handle_select_dir)
        self.view.load_button.config(command=self.handle_load_dir)
        self.view.result_tree.bind("<<TreeviewOpen>>", self.handle_tree_expand)
        self.view.result_tree.bind("<Double-1>", self.handle_file_open)

    def handle_file_open(self, event):
        """
        Summary:
            Treeviewのアイテムがダブルクリックされた際、ファイルであればOSの標準アプリで開きます。
        Args:
            event: Event - tkinterのイベントオブジェクト。
        UserAction:
            ファイルアイテムをダブルクリックした時 - OS標準アプリで対象ファイルを開く。
        """
        item_id = self.view.result_tree.focus()
        if not item_id:
            return

        values = self.view.result_tree.item(item_id, "values")
        if values and values[0] == "File":
            import os
            import shutil
            from datetime import datetime
            
            try:
                app_root = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
                temp_dir = os.path.join(app_root, "temp")
                os.makedirs(temp_dir, exist_ok=True)
                
                filename, ext = os.path.splitext(os.path.basename(item_id))
                ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
                new_filename = f"{filename}_{ts}{ext}"
                temp_file_path = os.path.join(temp_dir, new_filename)
                
                shutil.copy2(item_id, temp_file_path)
                os.startfile(temp_file_path)
            except Exception as e:
                self.view.set_status_message(f"ファイルを開けませんでした: {str(e)}")

    def handle_tree_expand(self, event):
        """
        Summary:
            Treeviewのノードが展開された際、未読み込みであればロジックを呼び出して配下を取得・追加します。
        Args:
            event: Event - tkinterのイベントオブジェクト。
        UserAction:
            ディレクトリ横の展開アイコン(+)をクリックした時 - 配下のアイテムを動的に読み込み、ツリーに追加して表示する。
        """
        item_id = self.view.result_tree.focus()
        if not item_id:
            return

        children = self.view.result_tree.get_children(item_id)
        if len(children) == 1 and children[0].endswith("|dummy"):
            import fws_apps.tkinter.fws_dir_viewer.views.dto.fws_dir_info_from_view_dto as fws_dir_info_from_view_dto
            dto = fws_dir_info_from_view_dto.FwsDirInfoFromViewDTO(dir_path=item_id)
            result_dto = self.logic.load_directory(dto)
            self.view.append_dir_items(item_id, result_dto)

    def handle_select_dir(self):
        """
        Summary:
            ディレクトリ選択ボタンをクリックした際、フォルダ選択ダイアログを表示してパスを入力欄に設定します。
        UserAction:
            参照ボタンをクリックした時 - フォルダ選択ダイアログを表示し、選択したパスを入力テキストボックスに設定する。
        """
        dir_path = filedialog.askdirectory(title="ディレクトリを選択してください")
        if not dir_path:
            return

        # パスを入力エリアにセットして自動読み込み
        self.view.set_dir_path(dir_path)
        self.handle_load_dir()

    def handle_load_dir(self):
        """
        Summary:
            読み込みボタンがクリックされた際、入力されたパスのディレクトリ情報を読み込みます。
        UserAction:
            読み込みボタンをクリックした時 - 入力されたパスの配下情報を取得し、グリッドに表示する。
        """
        input_dto = self.view.get_input_dto()
        result_dto = self.logic.load_directory(input_dto)
        self.view.set_dir_info(result_dto)
