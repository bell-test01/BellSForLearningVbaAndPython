"""
Summary:
    仕様書自動生成ツールアプリケーションのビジネスロジックモジュール。
"""

import os
import typing

import fws_apps.tkinter.fws_spec_generator.models.entity.fws_generator_entity as fws_generator_entity
import fws_apps.tkinter.fws_spec_generator.models.dto.fws_generator_from_view_dto as fws_generator_from_view_dto
import fws_apps.tkinter.fws_spec_generator.models.dto.fws_generator_to_view_dto as fws_generator_to_view_dto
import fws_apps.tkinter.fws_spec_generator.logics.fws_spec_parser_service as fws_spec_parser_service
import fws_apps.tkinter.fws_spec_generator.logics.fws_spec_generator_service as fws_spec_generator_service
import fws_lib.core.config.fws_config_manager as fws_config_manager


class FwsAppLogic:
    """
    Summary:
        仕様書自動生成ツールにおける永続化設定の管理、解析・生成プロセスの制御を行うクラス。
    """

    def __init__(
        self, test_config_manager=None, config_file_path: str = "data/config/config.json"
    ):
        self.config_manager = (
            test_config_manager
            if test_config_manager is not None
            else fws_config_manager.FwsConfigManager()
        )
        """FwsConfigManager - ファイル保存用設定マネージャ。"""

        self.config_file_path = config_file_path
        """str - 前回選択したパス設定を保存するJSONファイルのパス。"""

        self.parser_service = fws_spec_parser_service.FwsSpecParserService()
        """FwsSpecParserService - Pythonコードの構文木（AST）解析を行うサービス。"""

        self.generator_service = fws_spec_generator_service.FwsSpecGeneratorService()
        """FwsSpecGeneratorService - 仕様データをHTMLに整形出力するジェネレータサービス。"""

        # UI状態を保持するエンティティ
        self.entity = fws_generator_entity.FwsGeneratorEntity(
            id="",
            source_dir="",
            output_dir="",
            log_text="前回の設定をロードしました。\n",
            status_message="準備完了"
        )
        """FwsGeneratorEntity - UI表示および前回の設定状態を保持するエンティティ。"""

        self.load_config()

    def load_config(self) -> None:
        """
        Summary:
            設定ファイルからパス設定をロードします。
        """
        config = self.config_manager.load_json(self.config_file_path, default_value={})
        if isinstance(config, dict):
            source_dir = str(config.get("last_source_dir", ""))
            output_dir = str(config.get("last_output_dir", ""))
            
            self.entity.source_dir = source_dir
            self.entity.output_dir = output_dir

    def save_config(self) -> bool:
        """
        Summary:
            現在のパス設定を設定ファイルに保存します。
        Returns:
            bool - 保存成否。
        """
        config_data = {
            "last_source_dir": self.entity.source_dir,
            "last_output_dir": self.entity.output_dir
        }
        return self.config_manager.save_json(self.config_file_path, config_data)

    def update_paths(self, dto: fws_generator_from_view_dto.FwsGeneratorFromViewDTO) -> None:
        """
        Summary:
            画面から渡されたパスで Entity を更新し、設定を保存します。
        Args:
            dto: FwsGeneratorFromViewDTO - 画面から渡された入力DTO。
        """
        self.entity.source_dir = dto.source_dir
        self.entity.output_dir = dto.output_dir
        self.save_config()

    def execute_generation(self, append_log_callback: typing.Callable[[str], None]) -> str:
        """
        Summary:
            解析と HTML 生成を実行します。
        Args:
            append_log_callback: Callable[[str], None] - 実行ログを追記するためのコールバック関数。
        Returns:
            str - 生成されたHTMLファイルの絶対パス。
        Raises:
            Exception - 対象フォルダ内にPythonファイルが見つからなかった場合。
        """
        append_log_callback("--- 仕様書自動生成処理を開始します ---")
        append_log_callback(f"解析対象: {self.entity.source_dir}")
        append_log_callback(f"出力先: {self.entity.output_dir}")

        append_log_callback("Pythonソースファイルを探索し、ASTによる構文解析を実行中...")
        modules = self.parser_service.parse_directory(self.entity.source_dir)

        if not modules:
            append_log_callback("解析エラー: 対象フォルダ内に解析可能なPythonファイル（.py）が見つかりませんでした。")
            raise Exception("対象フォルダ内にPythonファイル（.py）が見つかりませんでした。")

        append_log_callback(f"解析完了。検出モジュール数: {len(modules)}")
        for m in modules:
            append_log_callback(f"  - {m.module_name} (クラス数: {len(m.classes)}, 関数数: {len(m.functions)})")

        append_log_callback("Javadoc風HTML仕様書を構築中...")
        output_file = os.path.join(self.entity.output_dir, "index.html")
        abs_path = self.generator_service.generate_html_spec(modules, output_file)

        append_log_callback("仕様書生成が成功しました。")
        append_log_callback(f"出力ファイル: {abs_path}")
        append_log_callback("--- 処理完了 ---")

        return abs_path

    def get_to_view_dto(self) -> fws_generator_to_view_dto.FwsGeneratorToViewDTO:
        """
        Summary:
            画面表示用のDTOオブジェクトを取得します。
        Returns:
            FwsGeneratorToViewDTO - 画面表示用DTO。
        """
        return self.entity.to_view_dto()
