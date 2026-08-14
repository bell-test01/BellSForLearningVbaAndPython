"""解析済"""
import dataclasses

@dataclasses.dataclass(frozen=True)
class FwsSpecVariableInfo:
    """
    Summary:
        クラスメンバ変数（属性）の情報を保持するエンティティクラス。
    """
    name:str
    type_name:str
    description:str