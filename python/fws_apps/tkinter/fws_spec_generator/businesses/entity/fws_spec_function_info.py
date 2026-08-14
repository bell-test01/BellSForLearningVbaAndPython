"""解析済"""
"""
Summary:
    Pythonソースコード解析結果のデータ構造を定義するエンティティモジュール。
"""
import dataclasses
import typing
from fws_apps.tkinter.fws_spec_generator.businesses.entity import fws_spec_parameter_info
from fws_apps.tkinter.fws_spec_generator.businesses.entity import fws_spec_exception_info
from fws_apps.tkinter.fws_spec_generator.businesses.entity import fws_spec_user_action_info
from fws_apps.tkinter.fws_spec_generator.businesses.entity import fws_spec_returns_info

@dataclasses.dataclass(frozen=True)
class FwsSpecFunctionInfo:
    """
    Summary:
        関数またはメソッドの解析情報を保持するエンティティクラス。
    """
    name:str
    summary:str
    description:str
    args:typing.List[fws_spec_parameter_info.FwsSpecParameterInfo]
    returns:typing.List[fws_spec_returns_info.FwsSpecReturnsInfo]
    raises: typing.List[fws_spec_exception_info.FwsSpecExceptionInfo]
    user_actions: typing.List[fws_spec_user_action_info.FwsSpecUserActionInfo]
    source_code: str
