"""
Summary:
    仕様書自動生成ツールにおける画面のイベント処理、対話制御、およびサービス連携を担当するモジュール。
"""

import os

import fws_apps.tkinter.fws_spec_generator.logics.fws_app_logic as fws_app_logic
import fws_apps.tkinter.fws_spec_generator.views.view.fws_app_view as fws_app_view
import fws_apps.tkinter.fws_spec_generator.models.dto.fws_generator_from_view_dto as fws_generator_from_view_dto


class FwsAppEvents:
    """
    Summary:
        メインウィンドウからのユーザーアクションおよびビジネスロジックのフロー制御を担当するクラス。
    """

    def __init__(self, test_view=None, test_logic=None):
        self.view = (
            test_view if test_view is not None else fws_app_view.FwsAppView()
        )
        self.logic = (
            test_logic
            if test_logic is not None
            else fws_app_logic.FwsAppLogic()
        )

        self._setup_event_bindings()
        self._initialize_app_data()

    def _setup_event_bindings(self):
        """
        Summary:
            各種ボタンに対して、Direct Event Binding でイベントハンドラーをバインドします。
        """
        self.view.btn_src.config(command=self.handle_select_source)
        self.view.btn_out.config(command=self.handle_select_output)
        self.view.btn_generate.config(command=self.handle_generate)

    def _initialize_app_data(self):
        """
        Summary:
            初期起動時にデータをロードし、ビューにデータを流し込みします。
        """
        self._refresh_view()
        self.view.set_status_message("データを読み込みました。")

    def handle_select_source(self):
        """
        Summary:
            解析対象フォルダの参照ボタンがクリックされたときに呼び出される処理。
        UserAction:
            解析対象フォルダの「参照...」ボタンをクリックした時 - フォルダ選択ダイアログを表示し、解析対象のPythonソースフォルダを指定する。
        """
        current_dto = self.view.get_input_dto()
        selected_dir = self.view.ask_directory(
            title="解析対象のPythonソースフォルダを選択してください",
            initialdir=current_dto.source_dir if current_dto.source_dir else None
        )
        if selected_dir:
            self.logic.entity.source_dir = selected_dir
            self.logic.entity.log_text += f"解析対象を選択しました: {selected_dir}\n"
            self.logic.entity.status_message = "解析対象フォルダ設定完了"
            self.logic.save_config()
            self._refresh_view()

    def handle_select_output(self):
        """
        Summary:
            仕様書出力先の参照ボタンがクリックされたときに呼び出される処理。
        UserAction:
            仕様書出力先の「参照...」ボタンをクリックした時 - フォルダ選択ダイアログを表示し、生成されたHTML仕様書の保存先フォルダを指定する。
        """
        current_dto = self.view.get_input_dto()
        selected_dir = self.view.ask_directory(
            title="仕様書HTMLの出力先フォルダを選択してください",
            initialdir=current_dto.output_dir if current_dto.output_dir else None
        )
        if selected_dir:
            self.logic.entity.output_dir = selected_dir
            self.logic.entity.log_text += f"出力先を選択しました: {selected_dir}\n"
            self.logic.entity.status_message = "出力先フォルダ設定完了"
            self.logic.save_config()
            self._refresh_view()

    def handle_generate(self):
        """
        Summary:
            仕様書自動生成ボタンがクリックされたときに呼び出される処理。
        UserAction:
            仕様書を自動生成するボタンをクリックした時 - 指定されたフォルダ内のPythonファイルを解析し、HTML仕様書ドキュメントを生成する。
        """
        dto = self.view.get_input_dto()

        if not dto.source_dir or not os.path.exists(dto.source_dir):
            self.view.show_warning("警告", "有効な解析対象フォルダを選択してください。")
            return

        if not dto.output_dir:
            self.view.show_warning("警告", "有効な仕様書出力先フォルダを選択してください。")
            return

        self.logic.update_paths(dto)

        self.view.set_generate_button_enabled(False)
        self.view.set_status_message("解析処理を実行中...")
        self.view.clear_log()

        try:
            abs_path = self.logic.execute_generation(append_log_callback=self.view.append_log)
            self.view.set_status_message("仕様書生成完了")
            self.view.show_info("生成完了", f"仕様書の生成が完了しました。\n\n出力先:\n{abs_path}")
        except Exception as e:
            err_msg = f"エラーが発生しました: {str(e)}"
            self.view.append_log(err_msg)
            self.view.set_status_message("生成エラー発生")
            self.view.show_error("エラー", err_msg)
        finally:
            self.view.set_generate_button_enabled(True)

    def _refresh_view(self):
        """
        Summary:
            ビジネスロジック層から最新の表示用DTOを取得し、ビューを再描画します。
        """
        to_view = self.logic.get_to_view_dto()
        self.view.update_view(to_view)
