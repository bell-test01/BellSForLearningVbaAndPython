import os
import tkinter as tk
import tkinter.ttk as ttk
import tkinter.filedialog as filedialog
import tkinter.messagebox as messagebox

import fws_apps.tkinter.fws_spec_generator.models.dto.fws_generator_from_view_dto as fws_generator_from_view_dto
import fws_apps.tkinter.fws_spec_generator.models.dto.fws_generator_to_view_dto as fws_generator_to_view_dto


class FwsAppView(tk.Tk):
    """仕様書生成画面の UI 表示層"""

    def __init__(self) -> None:
        """
        Summary:
            FwsAppViewクラスのインスタンスを初期化し、UIレイアウトを構築します。
        """
        super().__init__()
        self.title("Python仕様書自動生成ツール (Javadoc Style)")
        self.geometry("700x520")
        self.resizable(False, False)

        self._var_source_dir: tk.StringVar = tk.StringVar()
        self._var_output_dir: tk.StringVar = tk.StringVar()

        # Direct Event Binding 用の公開属性
        self.btn_src = None
        self.btn_out = None
        self.btn_generate = None
        self.entry_src = None
        self.entry_out = None
        self.txt_log = None
        self.lbl_status = None

        self._setup_ui()

    def _setup_ui(self) -> None:
        # パス選択フォームフレーム
        path_frame: ttk.LabelFrame = ttk.LabelFrame(self, text="パス設定", padding=15)
        path_frame.pack(fill=tk.X, padx=20, pady=15)

        # 解析対象フォルダ
        lbl_src: ttk.Label = ttk.Label(path_frame, text="解析対象フォルダ:")
        lbl_src.grid(row=0, column=0, sticky=tk.W, pady=5)

        self.entry_src = ttk.Entry(
            path_frame,
            textvariable=self._var_source_dir,
            font=("Segoe UI", 10),
        )
        self.entry_src.grid(row=0, column=1, sticky=tk.EW, padx=8, pady=5)

        self.btn_src = ttk.Button(
            path_frame,
            text="参照...",
        )
        self.btn_src.grid(row=0, column=2, sticky=tk.E, pady=5)

        # 仕様書出力先フォルダ
        lbl_out: ttk.Label = ttk.Label(path_frame, text="仕様書出力先:")
        lbl_out.grid(row=1, column=0, sticky=tk.W, pady=5)

        self.entry_out = ttk.Entry(
            path_frame,
            textvariable=self._var_output_dir,
            font=("Segoe UI", 10),
        )
        self.entry_out.grid(row=1, column=1, sticky=tk.EW, padx=8, pady=5)

        self.btn_out = ttk.Button(
            path_frame,
            text="参照...",
        )
        self.btn_out.grid(row=1, column=2, sticky=tk.E, pady=5)

        path_frame.columnconfigure(1, weight=1)

        # 生成実行ボタンフレーム
        action_frame: ttk.Frame = ttk.Frame(self)
        action_frame.pack(fill=tk.X, padx=20, pady=5)

        self.btn_generate = ttk.Button(
            action_frame,
            text="仕様書を自動生成する",
        )
        self.btn_generate.pack(side=tk.RIGHT, pady=5)

        # ログ表示エリアフレーム
        log_frame: ttk.LabelFrame = ttk.LabelFrame(self, text="実行ログ", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        scrollbar: ttk.Scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.txt_log = tk.Text(
            log_frame,
            font=("Consolas", 10),
            bd=1,
            relief=tk.SOLID,
            yscrollcommand=scrollbar.set
        )
        self.txt_log.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.txt_log.yview)

        # ステータスバーフレーム
        status_frame: ttk.Frame = ttk.Frame(self, height=24)
        status_frame.pack(fill=tk.X, side=tk.BOTTOM)

        self.lbl_status = ttk.Label(
            status_frame,
            text="準備完了",
            font=("Segoe UI", 9)
        )
        self.lbl_status.pack(side=tk.LEFT, padx=10, pady=2)

    # ==========================================
    # 明示的な Getter / Setter メソッド群
    # ==========================================

    def get_input_dto(self) -> fws_generator_from_view_dto.FwsGeneratorFromViewDTO:
        """画面入力値を FwsGeneratorFromViewDTO にまとめて取得"""
        return fws_generator_from_view_dto.FwsGeneratorFromViewDTO(
            source_dir=self._var_source_dir.get().strip(),
            output_dir=self._var_output_dir.get().strip(),
        )

    def update_view(self, state: fws_generator_to_view_dto.FwsGeneratorToViewDTO) -> None:
        """指定された FwsGeneratorToViewDTO に基づいて画面を一括更新"""
        self._var_source_dir.set(state.source_dir)
        self._var_output_dir.set(state.output_dir)
        
        self.txt_log.config(state=tk.NORMAL)
        self.txt_log.delete("1.0", tk.END)
        self.txt_log.insert(tk.END, state.log_text)
        self.txt_log.config(state=tk.DISABLED)

        self.lbl_status.config(text=state.status_message)

    def append_log(self, text: str) -> None:
        """ログ表示テキストエリアに文字列を追加"""
        self.txt_log.config(state=tk.NORMAL)
        self.txt_log.insert(tk.END, text + "\n")
        self.txt_log.see(tk.END)
        self.txt_log.config(state=tk.DISABLED)

    def clear_log(self) -> None:
        """ログテキストエリアをクリア"""
        self.txt_log.config(state=tk.NORMAL)
        self.txt_log.delete("1.0", tk.END)
        self.txt_log.config(state=tk.DISABLED)

    def set_generate_button_enabled(self, enabled: bool) -> None:
        """生成ボタンの活性状態を切り替え"""
        if enabled:
            self.btn_generate.config(state=tk.NORMAL)
        else:
            self.btn_generate.config(state=tk.DISABLED)

    def set_status_message(self, text: str) -> None:
        """ステータスメッセージ更新"""
        self.lbl_status.config(text=text)

    # ==========================================
    # ダイアログ・インタラクション用メソッド
    # ==========================================

    def ask_directory(self, title: str, initialdir: str = None) -> str:
        """フォルダ選択ダイアログを表示"""
        selected = filedialog.askdirectory(title=title, initialdir=initialdir if initialdir else None)
        return os.path.normpath(selected) if selected else ""

    def show_warning(self, title: str, message: str) -> None:
        """警告メッセージ"""
        messagebox.showwarning(title, message)

    def show_error(self, title: str, message: str) -> None:
        """エラーメッセージ"""
        messagebox.showerror(title, message)

    def show_info(self, title: str, message: str) -> None:
        """情報メッセージ"""
        messagebox.showinfo(title, message)
