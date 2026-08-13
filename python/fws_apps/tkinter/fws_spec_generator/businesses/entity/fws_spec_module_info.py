"""
Summary:
    Pythonソースコード解析結果のデータ構造を定義するエンティティモジュール。
"""
import dataclasses
import typing
from fws_apps.tkinter.fws_spec_generator.businesses.entity import fws_spec_class_info
from fws_apps.tkinter.fws_spec_generator.businesses.entity import fws_spec_function_info

@dataclasses.dataclass(frozen=True)
class FwsSpecModuleInfo:
    """
    Summary:
        ファイル（モジュール）全体の解析情報を保持するエンティティクラス。
    """
    file_path:str
    relative_path:str
    module_name:str
    summary:str
    description:str
    classes:typing.List[fws_spec_class_info.FwsSpecClassInfo]
    functions:typing.List[fws_spec_function_info.FwsSpecFunctionInfo]
    screen_name:typing.Optional[str] = None
    attachment_path:typing.Optional[str] = None