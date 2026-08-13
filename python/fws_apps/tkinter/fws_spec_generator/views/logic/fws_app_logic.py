"""
Summary:
    仕様書生成画面のロジック層。
Description:
    仕様書生成画面のビジネスロジックを定義する。
    イベント層から呼び出され、ビジネスロジックを実行し結果を返却する。
ScreenName:
    仕様書生成画面
"""

import os
import typing

from fws_lib.tkinter.file_util import fws_file_util
from fws_lib.tkinter.file_util import fws_file_dialog
from fws_apps.tkinter.fws_spec_generator.businesses.entity import fws_generator_entity
from fws_apps.tkinter.fws_spec_generator.views.view.dto import fws_generator_from_view_dto
from fws_apps.tkinter.fws_spec_generator.views.view.dto import fws_generator_to_view_dto
from fws_apps.tkinter.fws_spec_generator.businesses.business import fws_spec_parser_business
from fws_apps.tkinter.fws_spec_generator.businesses.business import fws_spec_generator_business

class FwsAppLogic:
    """
    Summary:
        仕様書生成画面のロジッククラス
    Description:
        仕様書生成画面のビジネスロジックを定義する。
    """
    #==========================================================================
    #コンストラクタ
    #==========================================================================
    def __init__(self):
        """
        Summary:
            コンストラクタ
        Description:
            ロジッククラスの初期設定を実施する。
            ビジネスロジックの初期設定を行う。
        Args:
            self: Any - FwsAppLogicクラスのインスタンス
        Returns:
            None - 無し
        """
        self.parser_business = fws_spec_parser_business.FwsSpecParserBusiness()
        """FwsSpecParserBusiness - Pythonコードの構文木（AST）解析を行うサービス。"""

        self.generator_business = fws_spec_generator_business.FwsSpecGeneratorBusiness()
        """FwsSpecGeneratorBusiness - 仕様データをHTMLに整形出力するジェネレータサービス。"""

        # UI状態を保持するエンティティ
        self.entity = fws_generator_entity.FwsGeneratorEntity(
            id="",
            source_dir="",
            output_dir="",
            log_text="",
            status_message="準備完了"
        )
        """FwsGeneratorEntity - UI表示および前回の設定状態を保持するエンティティ。"""

    #==========================================================================
    #パブリックメソッド
    #==========================================================================
    def select_analyzed_dir_path(self, initialdir: str) -> str:
        """
        Summary:
            解析対象フォルダのパスを選択する。
        Args:
            initialdir: str - 初期表示するディレクトリ。
        Returns:
            str - 解析対象フォルダのパス。
        """
        return self._selected_dir_path("解析対象のPythonソースフォルダを選択してください",initialdir)

    def select_output_dir_path(self, initialdir: str) -> str:
        """
        Summary:    
            仕様書出力先のフォルダのパスを選択する。
        Args:
            initialdir: str - 初期表示するディレクトリ。
        Returns:
            str - 仕様書出力先のフォルダのパス。
        """
        return self._selected_dir_path("仕様書出力先のフォルダを選択してください",initialdir)



    def execute_generation(self,analyzed_dir_path:str,output_dir_path:str ,set_txt_opelate_log_callback: typing.Callable[[str], None]) -> None:
        """
        Summary:
            解析と HTML 生成を実行します。
        Args:
            set_txt_opelate_log_callback: Callable[[str], None] - 実行ログを追記するためのコールバック関数。
        Returns:
            str - 生成されたHTMLファイルの絶対パス。
        """
        set_txt_opelate_log_callback("--- 仕様書自動生成処理を開始します ---")
        set_txt_opelate_log_callback(f"解析対象: {analyzed_dir_path}")
        set_txt_opelate_log_callback("Pythonソースファイルを探索し、ASTによる構文解析を実行中...")
        modules = self.parser_business.parse_directory(analyzed_dir_path)

        if not modules:
            set_txt_opelate_log_callback("解析エラー: 対象フォルダ内に解析可能なPythonファイル（.py）が見つかりませんでした。")
            return None

        set_txt_opelate_log_callback(f"解析完了。検出モジュール数: {len(modules)}")
        for m in modules:
            set_txt_opelate_log_callback(f"  - {m.module_name} (クラス数: {len(m.classes)}, 関数数: {len(m.functions)})")

        set_txt_opelate_log_callback("Javadoc風HTML仕様書を構築中...")
        output_file = os.path.join(output_dir_path, "index.html")
        abs_path = self.generator_business.generate_html_spec(modules, output_file)

        set_txt_opelate_log_callback(f"出力先: {output_dir_path}")
        set_txt_opelate_log_callback(f"フォルダ作成中")
        fws_file_util.create_directory(output_dir_path)
        set_txt_opelate_log_callback(f"フォルダ作成完了")


        set_txt_opelate_log_callback("仕様書生成が成功しました。")
        set_txt_opelate_log_callback(f"出力ファイル: {abs_path}")
        
        set_txt_opelate_log_callback("--- 処理完了 ---")
        

        return None









    def get_to_view_dto(self) -> fws_generator_to_view_dto.FwsGeneratorToViewDTO:
        """
        Summary:
            画面表示用のDTOオブジェクトを取得します。
        Returns:
            FwsGeneratorToViewDTO - 画面表示用DTO。
        """
        return self.entity.to_view_dto()


    #==========================================================================
    #プライベートメソッド
    #==========================================================================
    def _selected_dir_path(self,title:str,initialdir:str)->str:
        """
        Summary:    
            指定したタイトルのフォルダパスを選択する。
        Args:
            title: str - ダイアログのタイトル。
            initialdir: str - 初期表示するディレクトリ。
        """
        return fws_file_dialog.ask_directory(
            title=title,
            initialdir=initialdir
        )