"""解析済"""
import dataclasses

@dataclasses.dataclass(frozen=True)
class FwsSpecExceptionInfo:
    """
    Summary:
        発生しうる例外の情報を保持するエンティティクラス。
    """
    name:str
    description:str