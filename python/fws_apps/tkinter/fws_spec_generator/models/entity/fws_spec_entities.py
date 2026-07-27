"""
Summary:
    Pythonソースコード解析結果のデータ構造を定義するエンティティモジュール。
"""

import dataclasses
import typing


@dataclasses.dataclass(frozen=True)
class FwsVariableInfo:
    """
    Summary:
        クラスメンバ変数（属性）の情報を保持するエンティティクラス。
    """

    name: str
    type_name: str
    description: str


@dataclasses.dataclass(frozen=True)
class FwsParameterInfo:
    """
    Summary:
        関数またはメソッドの引数情報を保持するエンティティクラス。
    """

    name: str
    type_name: str
    description: str


@dataclasses.dataclass(frozen=True)
class FwsExceptionInfo:
    """
    Summary:
        発生しうる例外の情報を保持するエンティティクラス。
    """

    name: str
    description: str


@dataclasses.dataclass(frozen=True)
class FwsUserActionInfo:
    """
    Summary:
        エンジニア以外の読者に向けた画面操作仕様情報を保持するエンティティクラス。
    """

    trigger: str
    action: str


@dataclasses.dataclass(frozen=True)
class FwsFunctionInfo:
    """
    Summary:
        関数またはメソッドの解析情報を保持するエンティティクラス。
    """

    name: str
    summary: str
    description: str
    args: typing.List[FwsParameterInfo]
    returns_type: str
    returns_desc: str
    raises: typing.List[FwsExceptionInfo]
    user_actions: typing.List[FwsUserActionInfo]
    source_code: str


@dataclasses.dataclass(frozen=True)
class FwsClassInfo:
    """
    Summary:
        クラスの解析情報を保持するエンティティクラス。
    """

    name: str
    summary: str
    description: str
    attributes: typing.List[FwsVariableInfo]
    methods: typing.List[FwsFunctionInfo]


@dataclasses.dataclass(frozen=True)
class FwsModuleInfo:
    """
    Summary:
        ファイル（モジュール）全体の解析情報を保持するエンティティクラス。
    """

    file_path: str
    relative_path: str
    module_name: str
    summary: str
    description: str
    classes: typing.List[FwsClassInfo]
    functions: typing.List[FwsFunctionInfo]
    screen_name: typing.Optional[str] = None
    attachment_path: typing.Optional[str] = None
