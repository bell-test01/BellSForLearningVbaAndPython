"""
Summary:
    仕様書生成画面のUI表示層。
Description:
    仕様書生成画面の画面構成を定義する。
ScreenName:
    仕様書生成画面
"""

import tkinter as tk
import tkinter.ttk as ttk
from fws_apps.tkinter.fws_spec_generator.views.view.dto import fws_generator_from_view_dto
from fws_apps.tkinter.fws_spec_generator.views.view.dto import fws_generator_to_view_dto

class FwsAppView(tk.Tk):
    """
    Summary:
        仕様書生成画面の画面クラス
    Description:
        仕様書生成画面の画面構成を定義する。
    """

    #==========================================================================
    #コンストラクタ
    #==========================================================================
    def __init__(self) -> None:
        """
        Summary:
            コンストラクタ
        Description:
            画面クラスの初期設定を実施する。
            画面全体の初期設定、各種ウィジェットの定義指定、画面配置ファンクションの呼び出しを実施する。
        Args:
            self: Any - FwsAppViewクラスのインスタンス
        Returns:
            None - 無し
        """
        #画面全体の初期設定
        super().__init__()
        self.title("Python仕様書自動生成ツール (Javadoc Style)")
        self.geometry("700x520")
        self.resizable(False, False)

        #各種ウィジェットの定義
        self._ent_analyzed_dir_path: ttk.Entry = None
        """ttk.Entry - 解析対象フォルダの入力値"""
        self._ent_analyzed_dir_path_var: tk.StringVar = tk.StringVar()
        """tk.StringVar - 解析対象フォルダの入力値の変数"""
        self.btn_analyzed_dir_path_dialog: ttk.Button = None
        """ttk.Button - 解析対象フォルダ参照ダイアログ"""

        self._ent_output_dir_path: ttk.Entry = None
        """ttk.Entry - 出力先フォルダの入力値"""
        self._ent_output_dir_path_var: tk.StringVar = tk.StringVar()
        """tk.StringVar - 出力先フォルダの入力値の変数"""
        self.btn_output_dir_path_dialog: ttk.Button = None
        """ttk.Button - 出力先フォルダ参照ダイアログ"""
        
        self.btn_generate_spec: ttk.Button = None
        """ttk.Button - 仕様書生成ボタン"""
        
        self._txt_opelate_log: tk.Text = None
        """tk.Text - 実行ログ"""

        self._lbl_opelate_status: ttk.Label = None
        """ttk.Label - 実行ステータス"""
        self._lbl_opelate_status_var: tk.StringVar = tk.StringVar()
        """tk.StringVar - 実行ステータスの変数"""

        #画面配置ファンクションの呼び出し
        self._setup_ui()

    #==========================================================================
    #ゲッターセッター
    #==========================================================================
    def get_ent_analyzed_dir_path(self)->str:
        """
        Summary:
            解析対象フォルダの入力値を取得します。
        Args:
            self: Any - FwsAppViewクラスのインスタンス
        Returns:
            str - 解析対象フォルダの入力値
        """
        return self._ent_analyzed_dir_path_var.get()

    def set_ent_analyzed_dir_path(self,value:str)->None:
        """
        Summary:
            解析対象フォルダの入力値を設定します。
        Args:
            self: Any - FwsAppViewクラスのインスタンス
            value: str - 解析対象フォルダの入力値
        Returns:
            None - 無し
        """
        self._ent_analyzed_dir_path_var.set(value)

    def get_ent_output_dir_path(self)->str:
        """
        Summary:
            出力先フォルダの入力値を取得します。
        Args:
            self: Any - FwsAppViewクラスのインスタンス
        Returns:
            str - 出力先フォルダの入力値
        """
        return self._ent_output_dir_path_var.get()

    def set_ent_output_dir_path(self,value:str)->None:
        """
        Summary:
            出力先フォルダの入力値を設定します。
        Args:
            self: Any - FwsAppViewクラスのインスタンス
            value: str - 出力先フォルダの入力値
        Returns:
            None - 無し
        """
        self._ent_output_dir_path_var.set(value)

    def get_txt_opelate_log(self)->str:
        """
        Summary:
            実行ログの入力値を取得します。
        Args:
            self: Any - FwsAppViewクラスのインスタンス
        Returns:
            str - 実行ログの入力値
        """
        return self._txt_opelate_log.get("1.0", tk.END)
    
    def set_txt_opelate_log(self,value:str,is_append:bool=False)->None:
        """
        Summary:
            実行ログの入力値を設定します。
        Args:
            self: Any - FwsAppViewクラスのインスタンス
            value: str - 実行ログの入力値
            is_append: bool - 実行ログの入力値を追記するかどうか
        Returns:
            None - 無し
        """
        if not is_append:
            self._clear_txt_opelate_log()
        self._txt_opelate_log.insert(tk.END, value)
    
    def get_lbl_opelate_status(self)->str:
        """
        Summary:
            実行ステータスの入力値を取得します。
        Args:
            self: Any - FwsAppViewクラスのインスタンス
        Returns:
            str - 実行ステータスの入力値
        """
        return self._lbl_opelate_status_var.get()
    
    def set_lbl_opelate_status(self,value:str)->None:
        """
        Summary:
            実行ステータスの入力値を設定します。
        Args:
            self: Any - FwsAppViewクラスのインスタンス
            value: str - 実行ステータスの入力値
        Returns:
            None - 無し
        """
        self._lbl_opelate_status_var.set(value)

    #==========================================================================
    #プライベートメソッド
    #==========================================================================
    def _setup_ui(self) -> None:
        """
        Summary:
            画面のUIをセットアップします。
        Args:
            self: Any - FwsAppViewクラスのインスタンス
        Returns:
            None - 無し
        """
        #----------------------------------------------------
        # パス選択フォーム
        #----------------------------------------------------
        #フレーム
        _frm_path: ttk.LabelFrame = ttk.LabelFrame(self, text="パス設定", padding=15)
        _frm_path.pack(fill=tk.X, padx=20, pady=15)

        # 解析対象フォルダ
        _lbl_analyzed_dir_path: ttk.Label = ttk.Label(_frm_path, text="解析対象フォルダ:")
        _lbl_analyzed_dir_path.grid(row=0, column=0, sticky=tk.W, pady=5)

        self._ent_analyzed_dir_path = ttk.Entry(_frm_path,textvariable=self._ent_analyzed_dir_path_var,font=("Segoe UI", 10))
        self._ent_analyzed_dir_path.grid(row=0, column=1, sticky=tk.EW, padx=8, pady=5)

        self.btn_analyzed_dir_path_dialog = ttk.Button(_frm_path,text="参照...")
        self.btn_analyzed_dir_path_dialog.grid(row=0, column=2, sticky=tk.E, pady=5)

        # 仕様書出力先フォルダ
        _lbl_output_dir_path: ttk.Label = ttk.Label(_frm_path, text="仕様書出力先:")
        _lbl_output_dir_path.grid(row=1, column=0, sticky=tk.W, pady=5)

        self._ent_output_dir_path = ttk.Entry(_frm_path,textvariable=self._ent_output_dir_path_var,font=("Segoe UI", 10))
        self._ent_output_dir_path.grid(row=1, column=1, sticky=tk.EW, padx=8, pady=5)

        self.btn_output_dir_path_dialog = ttk.Button(_frm_path,text="参照...")
        self.btn_output_dir_path_dialog.grid(row=1, column=2, sticky=tk.E, pady=5)

        _frm_path.columnconfigure(1, weight=1)

        #----------------------------------------------------
        #生成実行ボタンフレーム
        #----------------------------------------------------
        #フレーム
        _action_frame: ttk.Frame = ttk.Frame(self)
        _action_frame.pack(fill=tk.X, padx=20, pady=5)

        #生成ボタン
        self.btn_generate_spec: ttk.Button = ttk.Button(_action_frame,text="仕様書を自動生成する")
        self.btn_generate_spec.pack(side=tk.RIGHT, pady=5)

        #----------------------------------------------------
        #ログ表示エリアフレーム
        #----------------------------------------------------
        #フレーム
        _log_frame: ttk.LabelFrame = ttk.LabelFrame(self, text="実行ログ", padding=10)
        _log_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        _scrollbar: ttk.Scrollbar = ttk.Scrollbar(_log_frame, orient=tk.VERTICAL)
        _scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self._txt_opelate_log: tk.Text = tk.Text(_log_frame,font=("Consolas", 10),bd=1,relief=tk.SOLID,yscrollcommand=_scrollbar.set)
        self._txt_opelate_log.pack(fill=tk.BOTH, expand=True)
        _scrollbar.config(command=self._txt_opelate_log.yview)

        #----------------------------------------------------
        #ステータスバーフレーム
        #----------------------------------------------------
        _status_frame: ttk.Frame = ttk.Frame(self, height=24)
        _status_frame.pack(fill=tk.X, side=tk.BOTTOM)

        self._lbl_opelate_status: ttk.Label = ttk.Label(_status_frame,textvariable=self._lbl_opelate_status_var,font=("Segoe UI", 9))
        self.set_lbl_opelate_status("準備完了")
        self._lbl_opelate_status.pack(side=tk.LEFT, padx=10, pady=2)

    def _clear_txt_opelate_log(self)->None:
        self._txt_opelate_log.delete("1.0", tk.END)
