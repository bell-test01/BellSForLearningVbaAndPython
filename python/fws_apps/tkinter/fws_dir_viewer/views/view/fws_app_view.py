"""
Summary:
    ディレクトリビューアー画面のUIレイアウトおよび表示制御を担当するビューモジュール。
ScreenName: ディレクトリビューアー画面
"""

import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk

import fws_apps.tkinter.fws_dir_viewer.views.dto.fws_dir_info_from_view_dto as fws_dir_info_from_view_dto
import fws_apps.tkinter.fws_dir_viewer.views.dto.fws_dir_info_to_view_dto as fws_dir_info_to_view_dto


class FwsAppView(tk.Tk):
    """
    Summary:
        ディレクトリビューアーのメインウィンドウ表示およびウィジェットの配置を管理するビュークラス。
    """

    def __init__(self):
        super().__init__()
        self.title("Directory Viewer (FWS - DTO Architecture)")
        self.geometry("800x600")
        self.minsize(600, 400)

        self.option_add("*font", ("Meiryo", 9))
        style = ttk.Style()
        style.configure(".", font=("Meiryo", 9))
        style.configure("Treeview", font=("Meiryo", 9), rowheight=22)
        style.configure("Treeview.Heading", font=("Meiryo", 9, "bold"))

        self.dir_path_entry: ttk.Entry = None
        self.select_dir_button: ttk.Button = None
        self.load_button: ttk.Button = None
        self.result_tree: ttk.Treeview = None
        self.status_label: ttk.Label = None

        self._create_widgets()

    def _create_widgets(self):
        """
        Summary:
            ウィンドウ内にUIウィジェットを配置します。
        """
        main_frame = ttk.Frame(self, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 上部: パス入力エリア
        path_frame = ttk.LabelFrame(main_frame, text="対象ディレクトリ", padding=5)
        path_frame.pack(fill=tk.X, pady=(0, 10))

        self.dir_path_entry = ttk.Entry(path_frame, font=("Meiryo", 10))
        self.dir_path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        self.select_dir_button = ttk.Button(path_frame, text="参照...")
        self.select_dir_button.pack(side=tk.LEFT, padx=(0, 5))

        self.load_button = ttk.Button(path_frame, text="読み込み")
        self.load_button.pack(side=tk.LEFT)

        # 中央: ツリービュー
        result_frame = ttk.LabelFrame(main_frame, text="ディレクトリおよびファイル一覧", padding=5)
        result_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))

        self.result_tree = ttk.Treeview(result_frame, columns=("Type", "Size", "Modified"), show="tree headings")
        
        self.result_tree.heading("#0", text="名前", command=lambda: self._sort_tree_column("#0", False))
        self.result_tree.column("#0", width=350, anchor=tk.W)
        self.result_tree.heading("Type", text="種類", command=lambda: self._sort_tree_column("Type", False))
        self.result_tree.column("Type", width=80, anchor=tk.W)
        self.result_tree.heading("Size", text="サイズ (Bytes)", command=lambda: self._sort_tree_column("Size", False))
        self.result_tree.column("Size", width=120, anchor=tk.E)
        self.result_tree.heading("Modified", text="更新日時", command=lambda: self._sort_tree_column("Modified", False))
        self.result_tree.column("Modified", width=150, anchor=tk.W)

        result_scroll_y = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.result_tree.yview)
        result_scroll_x = ttk.Scrollbar(result_frame, orient=tk.HORIZONTAL, command=self.result_tree.xview)
        self.result_tree.configure(yscrollcommand=result_scroll_y.set, xscrollcommand=result_scroll_x.set)

        result_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        result_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.result_tree.pack(fill=tk.BOTH, expand=True)

        # 下部: ステータスバー
        self.status_label = ttk.Label(main_frame, text="ディレクトリパスを入力し、読み込みボタンを押してください。", relief=tk.SUNKEN, anchor=tk.W)
        self.status_label.pack(fill=tk.X, side=tk.BOTTOM)

    def get_input_dto(self) -> fws_dir_info_from_view_dto.FwsDirInfoFromViewDTO:
        """
        Summary:
            画面上のディレクトリパスをDTOとして取得します。
        Returns:
            FwsDirInfoFromViewDTO - 画面から取得した入力DTO。
        """
        dir_path = self.dir_path_entry.get().strip()
        return fws_dir_info_from_view_dto.FwsDirInfoFromViewDTO(dir_path=dir_path)

    def set_dir_path(self, dir_path: str):
        """
        Summary:
            入力欄にディレクトリパスを設定します。
        Args:
            dir_path: str - 設定するディレクトリパス。
        """
        self.dir_path_entry.delete(0, tk.END)
        self.dir_path_entry.insert(0, dir_path)

    def set_dir_info(self, result: fws_dir_info_to_view_dto.FwsDirInfoToViewDTO):
        """
        Summary:
            ルートディレクトリの結果DTOに基づいてTreeview全体を再描画し、ステータスを更新します。
        Args:
            result: FwsDirInfoToViewDTO - 実行結果DTO。
        """
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

        if result.is_success and result.items:
            self._insert_items("", result.items)

        self.set_status_message(result.message)

    def append_dir_items(self, parent_iid: str, result: fws_dir_info_to_view_dto.FwsDirInfoToViewDTO):
        """
        Summary:
            指定された親ノード配下のダミーノードを削除し、結果DTOに基づいてアイテムを追加します。
        Args:
            parent_iid: str - 親ノードの固有ID（絶対パス）。
            result: FwsDirInfoToViewDTO - 実行結果DTO。
        """
        for child in self.result_tree.get_children(parent_iid):
            self.result_tree.delete(child)

        if result.is_success and result.items:
            self._insert_items(parent_iid, result.items)

        self.set_status_message(result.message)

    def _insert_items(self, parent_iid: str, items: list[tuple]):
        """
        Summary:
            指定された親ノードの配下にアイテムリストを挿入します。ディレクトリの場合はダミーノードを追加します。
        Args:
            parent_iid: str - 親ノードのID。
            items: list[tuple] - (名前, 種類, サイズ, 最終更新日時, 絶対パス) のリスト。
        """
        for item in items:
            name, item_type, size, mtime, full_path = item
            node = self.result_tree.insert(parent_iid, tk.END, iid=full_path, text=name, values=(item_type, size, mtime))
            if item_type == "Dir":
                self.result_tree.insert(node, tk.END, iid=f"{full_path}|dummy", text="Loading...")

    def set_status_message(self, message: str):
        """
        Summary:
            ステータスバーのメッセージを更新します。
        Args:
            message: str - 表示するメッセージ文字列。
        """
        self.status_label.config(text=message)

    def _sort_tree_column(self, col: str, reverse: bool):
        """
        Summary:
            指定された列でTreeviewのアイテムを再帰的にソートします。
        Args:
            col: str - ソート対象の列ID（"#0", "Type", "Size", "Modified"）。
            reverse: bool - Trueなら降順、Falseなら昇順。
        """
        self._sort_children("", col, reverse)
        self.result_tree.heading(col, command=lambda: self._sort_tree_column(col, not reverse))

    def _sort_children(self, parent_iid: str, col: str, reverse: bool):
        """
        Summary:
            指定された親ノードの直下にあるアイテムをソートし、再帰的に配下もソートします。
        Args:
            parent_iid: str - 親ノードのID。
            col: str - ソート対象の列ID。
            reverse: bool - 降順かどうか。
        """
        children = list(self.result_tree.get_children(parent_iid))
        if not children:
            return

        if len(children) == 1 and children[0].endswith("|dummy"):
            return

        items_to_sort = []
        for child_id in children:
            if col == "#0":
                val = self.result_tree.item(child_id, "text")
            else:
                val = self.result_tree.set(child_id, col)
            
            sort_val = val
            if col == "Size":
                try:
                    sort_val = int(val)
                except ValueError:
                    sort_val = -1

            items_to_sort.append((sort_val, child_id))

        items_to_sort.sort(reverse=reverse, key=lambda x: x[0])

        for index, (_, child_id) in enumerate(items_to_sort):
            self.result_tree.move(child_id, parent_iid, index)
            self._sort_children(child_id, col, reverse)
