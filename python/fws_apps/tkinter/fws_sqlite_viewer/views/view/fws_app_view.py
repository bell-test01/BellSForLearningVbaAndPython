"""
Summary:
    SQLite Viewer画面のUIレイアウトおよび表示制御を担当するビューモジュール。
ScreenName: SQLite Viewer画面
"""

import re
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk

import fws_apps.tkinter.fws_sqlite_viewer.models.dto.fws_db_state_to_view_dto as fws_db_state_to_view_dto
import fws_apps.tkinter.fws_sqlite_viewer.models.dto.fws_query_from_view_dto as fws_query_from_view_dto
import fws_apps.tkinter.fws_sqlite_viewer.models.dto.fws_query_result_to_view_dto as fws_query_result_to_view_dto


class FwsSqlEditor(ttk.Frame):
    """
    Summary:
        行番号表示、簡易シンタックスハイライト、およびUndo/Redo機能を備えたSQLテキストエディタ。
    """

    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent)

        # 行番号用キャンバス
        self.line_canvas = tk.Canvas(self, width=40, bg="#f0f0f0", bd=0, highlightthickness=0)
        self.line_canvas.pack(side=tk.LEFT, fill=tk.Y)

        # テキスト入力エリア
        self.text_widget = tk.Text(
            self,
            font=("Meiryo", 9),
            undo=True,
            wrap=tk.NONE,
            *args,
            **kwargs
        )
        self.text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Y軸スクロールバー
        self.scroll_y = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.text_widget.yview)
        self.scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        self.text_widget.configure(yscrollcommand=self._on_scroll)

        # X軸スクロールバーはテキストエリアのスクロールと連動（必要に応じて）
        self.scroll_x = ttk.Scrollbar(parent, orient=tk.HORIZONTAL, command=self.text_widget.xview)
        # ※親要素側で管理しやすくするため、ここではY軸スクロールバーのみ統合

        # シンタックスハイライト用のタグ色設定
        self.text_widget.tag_config("keyword", foreground="#0000ff", font=("Meiryo", 9, "bold"))
        self.text_widget.tag_config("string", foreground="#a52a2a")
        self.text_widget.tag_config("comment", foreground="#008000")
        self.text_widget.tag_config("number", foreground="#800080")

        self.text_widget.bind("<KeyRelease>", self._on_key_release)
        self.text_widget.bind("<Configure>", lambda e: self.redraw_line_numbers())

        self.redraw_line_numbers()
        self.highlight_sql()

    def _on_scroll(self, *args):
        self.scroll_y.set(*args)
        self.redraw_line_numbers()

    def _on_key_release(self, event):
        self.redraw_line_numbers()
        self.highlight_sql()

    def get(self, *args, **kwargs):
        return self.text_widget.get(*args, **kwargs)

    def insert(self, *args, **kwargs):
        self.text_widget.insert(*args, **kwargs)
        self.redraw_line_numbers()
        self.highlight_sql()

    def delete(self, *args, **kwargs):
        self.text_widget.delete(*args, **kwargs)
        self.redraw_line_numbers()
        self.highlight_sql()

    def redraw_line_numbers(self):
        self.line_canvas.delete("all")
        i = self.text_widget.index("@0,0")
        while True:
            dline = self.text_widget.dlineinfo(i)
            if dline is None:
                break
            y = dline[1]
            linenum = i.split(".")[0]
            # 右寄せで行番号を描画
            self.line_canvas.create_text(
                35, y + 2, anchor="ne", text=linenum, fill="#888888", font=("Meiryo", 8)
            )
            i = self.text_widget.index(f"{i}+1line")

    def highlight_sql(self):
        self.text_widget.tag_remove("keyword", "1.0", tk.END)
        self.text_widget.tag_remove("string", "1.0", tk.END)
        self.text_widget.tag_remove("comment", "1.0", tk.END)
        self.text_widget.tag_remove("number", "1.0", tk.END)

        content = self.text_widget.get("1.0", tk.END)

        keywords = {
            "select", "from", "where", "join", "on", "group", "by", "order", "having",
            "insert", "into", "values", "update", "set", "delete", "create", "table",
            "drop", "alter", "index", "view", "and", "or", "not", "in", "is", "null",
            "like", "as", "left", "right", "inner", "outer", "cross", "sqlite_master"
        }

        # コメント (-- ...)
        for match in re.finditer(r"--.*", content):
            start = f"1.0 + {match.start()} chars"
            end = f"1.0 + {match.end()} chars"
            self.text_widget.tag_add("comment", start, end)

        # 文字列リテラル ('...' / "...")
        for match in re.finditer(r"'[^']*'|\"[^\"]*\"", content):
            start = f"1.0 + {match.start()} chars"
            end = f"1.0 + {match.end()} chars"
            self.text_widget.tag_add("string", start, end)

        # SQL予約語（単語境界）
        for match in re.finditer(r"\b[a-zA-Z_][a-zA-Z0-9_]*\b", content):
            word = match.group(0).lower()
            if word in keywords:
                start = f"1.0 + {match.start()} chars"
                end = f"1.0 + {match.end()} chars"
                self.text_widget.tag_add("keyword", start, end)

        # 数値
        for match in re.finditer(r"\b\d+\b", content):
            start = f"1.0 + {match.start()} chars"
            end = f"1.0 + {match.end()} chars"
            self.text_widget.tag_add("number", start, end)

        # ハイライトタグの重ね順を設定（コメント・文字列を優先）
        self.text_widget.tag_raise("comment")
        self.text_widget.tag_raise("string")
        self.text_widget.tag_raise("keyword")


class FwsAppView(tk.Tk):
    """
    Summary:
        SQLite Viewerのメインウィンドウ表示およびウィジェットの配置を管理するビュークラス。
    """


    def __init__(self):
        super().__init__()
        self.title("SQLite Viewer (FWS - DTO Architecture)")
        self.geometry("800x600")
        minimum_width: int = 700
        minimum_height: int = 500
        self.minsize(minimum_width, minimum_height)

        # 全体のフォントをメイリオの9ptに設定
        self.option_add("*font", ("Meiryo", 9))

        # ttkウィジェットのフォント設定
        style = ttk.Style()
        style.configure(".", font=("Meiryo", 9))
        style.configure("Treeview", font=("Meiryo", 9), rowheight=22)
        style.configure("Treeview.Heading", font=("Meiryo", 9, "bold"))

        # Direct Event Binding 用の公開属性の型注釈
        self.menu_bar: tk.Menu = None
        self.file_menu: tk.Menu = None
        self.db_path_entry: ttk.Entry = None
        self.select_db_button: ttk.Button = None
        self.reload_db_button: ttk.Button = None
        self.table_tree: ttk.Treeview = None
        self.sql_text: FwsSqlEditor = None
        self.execute_button: ttk.Button = None
        self.result_tree: ttk.Treeview = None
        self.status_label: ttk.Label = None

        self._create_widgets()

    def _create_widgets(self):
        """
        Summary:
            ウィンドウ内にUIウィジェットを配置します。
        """
        # --- メニューバーの作成 ---
        self.menu_bar = tk.Menu(self)
        self.config(menu=self.menu_bar)

        self.file_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="ファイル", menu=self.file_menu)
        self.file_menu.add_command(label="新規データベース作成...")
        self.file_menu.add_command(label="データベースを開く...")
        self.file_menu.add_command(label="再読み込み")
        self.file_menu.add_separator()
        self.file_menu.add_command(label="終了", command=self.quit)

        # --- メインフレーム ---
        main_frame = ttk.Frame(self, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- 上部: ファイル選択エリア ---
        db_frame = ttk.LabelFrame(main_frame, text="データベース選択", padding=5)
        db_frame.pack(fill=tk.X, pady=(0, 10))

        self.db_path_entry = ttk.Entry(db_frame, font=("Helvetica", 10))
        self.db_path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        self.select_db_button = ttk.Button(db_frame, text="ファイル参照...")
        self.select_db_button.pack(side=tk.RIGHT, padx=(5, 0))

        self.reload_db_button = ttk.Button(db_frame, text="再読み込み")
        self.reload_db_button.pack(side=tk.RIGHT)

        # --- 中央: 分割パネル ---
        paned = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, pady=(0, 5))

        # --- 左側: テーブル一覧エリア ---
        left_frame = ttk.LabelFrame(paned, text="テーブル一覧", padding=5, width=200)
        paned.add(left_frame, weight=1)

        self.table_tree = ttk.Treeview(
            left_frame, columns=("name",), show="headings", selectmode="browse"
        )
        self.table_tree.heading("name", text="テーブル名")
        self.table_tree.column("name", width=180, anchor=tk.W)

        left_scroll = ttk.Scrollbar(
            left_frame, orient=tk.VERTICAL, command=self.table_tree.yview
        )
        self.table_tree.configure(yscrollcommand=left_scroll.set)

        left_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.table_tree.pack(fill=tk.BOTH, expand=True)

        # --- 右側: SQL入力および実行結果エリア ---
        right_pane = ttk.PanedWindow(paned, orient=tk.VERTICAL)
        paned.add(right_pane, weight=3)

        # 右上: SQLクエリ入力
        sql_frame = ttk.LabelFrame(right_pane, text="SQLクエリ入力", padding=5, height=150)
        right_pane.add(sql_frame, weight=1)

        self.sql_text = FwsSqlEditor(sql_frame, height=5)
        self.sql_text.pack(fill=tk.BOTH, expand=True, pady=(0, 5))
        self.sql_text.insert(tk.END, "SELECT name FROM sqlite_master WHERE type='table';")

        self.execute_button = ttk.Button(sql_frame, text="SQLクエリ実行")
        self.execute_button.pack(anchor=tk.E)

        # 右下: 実行結果表示
        result_frame = ttk.LabelFrame(right_pane, text="実行結果", padding=5, height=250)
        right_pane.add(result_frame, weight=2)

        self.result_tree = ttk.Treeview(result_frame, show="headings")

        result_scroll_y = ttk.Scrollbar(
            result_frame, orient=tk.VERTICAL, command=self.result_tree.yview
        )
        result_scroll_x = ttk.Scrollbar(
            result_frame, orient=tk.HORIZONTAL, command=self.result_tree.xview
        )
        self.result_tree.configure(
            yscrollcommand=result_scroll_y.set, xscrollcommand=result_scroll_x.set
        )

        result_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        result_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.result_tree.pack(fill=tk.BOTH, expand=True)

        # --- 最下部: ステータスバー ---
        self.status_label = ttk.Label(
            main_frame, text="データベースファイルを開いてください。", relief=tk.SUNKEN, anchor=tk.W
        )
        self.status_label.pack(fill=tk.X, side=tk.BOTTOM)

    # ==========================================
    # 明示的な Getter / Setter メソッド群
    # ==========================================

    def get_input_dto(self) -> fws_query_from_view_dto.FwsQueryFromViewDTO:
        """
        Summary:
            画面上のデータベースパスとSQL入力クエリをFwsQueryFromViewDTOとして取得します。
        Returns:
            FwsQueryFromViewDTO - 画面から取得したクエリ実行入力DTO。
        """
        db_path = self.db_path_entry.get().strip()
        query = self.sql_text.get("1.0", tk.END).strip()
        return fws_query_from_view_dto.FwsQueryFromViewDTO(
            db_path=db_path,
            query=query,
        )

    def set_db_state(self, state: fws_db_state_to_view_dto.FwsDbStateToViewDTO):
        """
        Summary:
            データベース接続状態に基づいて、DBパスおよびテーブル一覧（Treeview）を再描画します。
        Args:
            state: FwsDbStateToViewDTO - データベース接続状態の表示用DTO。
        """
        # パス表示Entryを更新
        self.db_path_entry.delete(0, tk.END)
        self.db_path_entry.insert(0, state.db_path)

        # テーブルツリーを再描画
        for item in self.table_tree.get_children():
            self.table_tree.delete(item)

        for tbl in state.tables:
            self.table_tree.insert("", tk.END, iid=tbl, values=(tbl,))

    def set_query_result(
        self, result: fws_query_result_to_view_dto.FwsQueryResultToViewDTO
    ):
        """
        Summary:
            クエリ実行結果に基づいて結果Treeviewを再描画し、ステータスを更新します。
        Args:
            result: FwsQueryResultToViewDTO - 表示用クエリ実行結果DTO。
        """
        # Treeviewの項目と列定義を完全にクリア
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

        # 動的な列再設定
        cols = tuple(result.columns)
        self.result_tree.configure(columns=cols)

        # フォント幅測定用の設定
        measure_font = tkfont.Font(font=("Meiryo", 9))
        col_widths = {}

        for col in cols:
            # 見出し（ヘッダ）のテキスト長を初期値にする
            col_widths[col] = measure_font.measure(str(col)) + 25

        # データ行の挿入
        if result.is_success and result.rows:
            for row in result.rows:
                self.result_tree.insert("", tk.END, values=row)

                # 幅測定の実施
                for idx, val in enumerate(row):
                    if idx < len(cols):
                        col = cols[idx]
                        val_str = str(val) if val is not None else ""
                        
                        # セル内容に応じた幅測定（パディング分を加算）
                        val_width = measure_font.measure(val_str) + 20
                        if val_width > col_widths[col]:
                            col_widths[col] = val_width

        # 動的に測定した列幅の適用（すべて左寄せ）
        for col in cols:
            anchor_val = tk.W
            # 列幅の最低値・最高値を設定して極端なサイズを防ぐ
            col_width = max(60, min(500, col_widths[col]))
            
            self.result_tree.heading(col, text=col)
            self.result_tree.column(col, width=col_width, anchor=anchor_val, stretch=False)

        self.set_status_message(result.message)

    def get_selected_table(self) -> str | None:
        """
        Summary:
            テーブル一覧Treeviewで選択されているテーブル名を取得します。
        Returns:
            str | None - 選択されたテーブル名。選択されていない場合はNone。
        """
        selection = self.table_tree.selection()
        if selection:
            return selection[0]
        return None

    def set_sql_text(self, text: str):
        """
        Summary:
            SQLクエリ入力エリアのテキストを書き換えます。
        Args:
            text: str - 設定するSQLテキスト文字列。
        """
        self.sql_text.delete("1.0", tk.END)
        self.sql_text.insert(tk.END, text)

    def set_status_message(self, message: str):
        """
        Summary:
            画面下部のステータスバーのメッセージを更新します。
        Args:
            message: str - 表示するメッセージ文字列。
        """
        self.status_label.config(text=message)
