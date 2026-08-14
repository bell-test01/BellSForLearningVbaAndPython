"""解析済"""
"""
Summary:
    Pythonソースコード解析結果のデータ構造を定義するエンティティモジュール。
"""
import dataclasses
import typing

from fws_apps.tkinter.fws_spec_generator.businesses.entity import fws_spec_function_info
from fws_apps.tkinter.fws_spec_generator.businesses.entity import fws_spec_variable_info

@dataclasses.dataclass(frozen=True)
class FwsSpecClassInfo:
    """
    Summary:
        クラスの解析情報を保持するエンティティクラス。
    """
    name:str
    summary:str
    description:str
    attributes:typing.List[fws_spec_variable_info.FwsSpecVariableInfo]
    methods:typing.List[fws_spec_function_info.FwsSpecFunctionInfo]