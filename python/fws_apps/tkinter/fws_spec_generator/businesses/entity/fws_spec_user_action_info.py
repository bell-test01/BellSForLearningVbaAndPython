"""解析済"""
import dataclasses

@dataclasses.dataclass(frozen=True)
class FwsSpecUserActionInfo:
    """
    Summary:
        エンジニア以外の読者に向けた画面操作仕様情報を保持するエンティティクラス。
    """
    trigger:str
    action:str