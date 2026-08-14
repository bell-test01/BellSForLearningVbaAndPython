"""解析済"""
import dataclasses

@dataclasses.dataclass(frozen=True)
class FwsSpecReturnsInfo:
    """
    Summary:
        関数またはメソッドの戻り値情報を保持するエンティティクラス。
    """
    type_name:str
    description:str
