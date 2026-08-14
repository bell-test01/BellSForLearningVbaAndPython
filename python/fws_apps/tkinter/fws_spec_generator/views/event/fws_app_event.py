"""解析済"""
"""
Summary:
    仕様書生成画面のイベント層。
Description:
    仕様書生成画面のイベント処理、対話制御、およびサービス連携を担当する。
    ロジック層を呼び出し、ビジネスロジックを実行する。
    ビジネスロジックの結果をUIに反映させる。
ScreenName:
    仕様書生成画面
"""
import os
import traceback
from fws_lib.tkinter.exception_util import fws_exception_util
from fws_lib.tkinter.file_util import fws_file_util
from fws_apps.tkinter.fws_spec_generator.views.logic import fws_app_logic
from fws_apps.tkinter.fws_spec_generator.views.view import fws_app_view


class FwsAppEvent:
    """
    Summary:
        仕様書生成画面のイベントクラス
    Description:
        仕様書生成画面のイベント処理、対話制御、およびサービス連携を担当する。
        ロジック層を呼び出し、ビジネスロジックを実行する。
        ビジネスロジックの結果をUIに反映させる。
    """
    #==========================================================================
    #コンストラクタ
    #==========================================================================
    def __init__(self):
        """
        Summary:
            コンストラクタ
        Description:
            イベントクラスの初期設定を実施する。
            ビジネスロジックの初期設定を行う。
        Args:
            self: Any - FwsAppEventクラスのインスタンス
        Returns:
            None - 無し
        """
        self.fws_app_view_obj:fws_app_view.FwsAppView = fws_app_view.FwsAppView()
        """FwsAppView - 仕様書生成画面のビュー"""
        self.fws_app_logic_obj:fws_app_logic.FwsAppLogic = fws_app_logic.FwsAppLogic()
        """FwsAppLogic - 仕様書生成画面のロジック層"""
        
        self._initialize_app_data()
        self._setup_event_bindings()

    #==========================================================================
    #画面初期設定
    #==========================================================================
    def _initialize_app_data(self):
        """
        Summary:
            初期起動時にデータをロードし、ビューにデータを流し込みます。
        Args:
            self: Any - FwsAppEventクラスのインスタンス
        Returns:
            None - 無し
        """
        try:
            self.fws_app_view_obj.set_lbl_opelate_status("データを読み込みました。")
        except Exception as e:
            fws_exception_util.test_except(e)

    #==========================================================================
    #イベントバインド
    #==========================================================================
    def _setup_event_bindings(self):
        """
        Summary:
            各種ボタンに対して、Direct Event Binding でイベントハンドラーをバインドします。
        Args:
            self: Any - FwsAppEventクラスのインスタンス
        Returns:
            None - 無し
        """
        try:
            self.fws_app_view_obj.btn_analyzed_dir_path_dialog.config(command=self._btn_analyzed_dir_path_dialog_click)
            self.fws_app_view_obj.btn_output_dir_path_dialog.config(command=self._btn_output_dir_path_dialog_click)
            self.fws_app_view_obj.btn_generate_spec.config(command=self._btn_generate_spec_click)
        except Exception as e:
            fws_exception_util.test_except(e)

    #==========================================================================
    #イベントハンドラー
    #==========================================================================
    def _btn_analyzed_dir_path_dialog_click(self):
        """
        Summary:
            解析対象フォルダの参照ボタンがクリックされたときに呼び出される処理。
        Args:
            self: Any - FwsAppEventクラスのインスタンス
        Returns:
            None - 無し
        UserAction:
            解析対象フォルダの「参照...」ボタンをクリックした時 - フォルダ選択ダイアログを表示し、解析対象のPythonソースフォルダを指定する。指定結果は実行ログエリアに表示する。
        """
        try:
            selected_dir = self.fws_app_logic_obj.select_analyzed_dir_path(self.fws_app_view_obj.get_ent_analyzed_dir_path())
            if not selected_dir:
                return

            self.fws_app_view_obj.set_ent_analyzed_dir_path(selected_dir)
            self.fws_app_view_obj.set_txt_opelate_log(f"解析対象を選択しました: {selected_dir}\n",False)
            self.fws_app_view_obj.set_lbl_opelate_status("解析対象フォルダ設定完了")
        except Exception as e:
            fws_exception_util.test_except(e)


    def _btn_output_dir_path_dialog_click(self):
        """
        Summary:
            仕様書出力先の参照ボタンがクリックされたときに呼び出される処理。
        Args:
            self: Any - FwsAppEventクラスのインスタンス
        Returns:
            None - 無し
        UserAction:
            仕様書出力先の「参照...」ボタンをクリックした時 - フォルダ選択ダイアログを表示し、生成されたHTML仕様書の保存先フォルダを指定する。指定結果は実行ログエリアに表示する。
        """
        try:
            selected_dir = self.fws_app_logic_obj.select_output_dir_path(self.fws_app_view_obj.get_ent_output_dir_path())
            if not selected_dir:
                return

            self.fws_app_view_obj.set_ent_output_dir_path(selected_dir)
            self.fws_app_view_obj.set_txt_opelate_log(f"出力先を選択しました: {selected_dir}\n",False)
            self.fws_app_view_obj.set_lbl_opelate_status("出力先フォルダ設定完了")
        except Exception as e:
            fws_exception_util.test_except(e)

    def _btn_generate_spec_click(self):
        """
        Summary:
            仕様書自動生成ボタンがクリックされたときに呼び出される処理。
        Args:
            self: Any - FwsAppEventクラスのインスタンス
        Returns:
            None - 無し
        UserAction:
            仕様書を自動生成するボタンをクリックした時 - 指定されたフォルダ内のPythonファイルを解析し、HTML仕様書ドキュメントを生成する。
        """
        try:
            analyzed_dir_path = self.fws_app_view_obj.get_ent_analyzed_dir_path()
            if not fws_file_util.is_directory_exists(analyzed_dir_path):
                self.fws_app_view_obj.show_warning("警告", "有効な解析対象フォルダを選択してください。")
                return

            output_dir_path = self.fws_app_view_obj.get_ent_output_dir_path()
            if not fws_file_util.is_directory_exists(output_dir_path):
                self.fws_app_view_obj.show_warning("警告", "有効な仕様書出力先フォルダを選択してください。")
                return

            self.fws_app_logic_obj.execute_generation(analyzed_dir_path, output_dir_path,self.set_txt_opelate_log_callbck)

            self.fws_app_view_obj.set_txt_opelate_log(f"仕様書を生成しました: {output_dir_path}\n",True)
        except Exception as e:
            fws_exception_util.test_except(e)
            error_details: str = traceback.format_exc()
            print("=== エラーが発生しました ===")
            print(error_details)
            print("============================")

    #==========================================================================
    #コールバックメソッド
    #==========================================================================
    def set_txt_opelate_log_callbck(self, message:str)->None:
        """
        Summary:
            ログメッセージを表示するためのコールバックメソッド。
        Args:
            self: Any - FwsAppEventクラスのインスタンス
            message: str - 表示するログメッセージ
        Returns:
            None - 無し
        """
        self.fws_app_view_obj.set_txt_opelate_log(message + "\n",True)

