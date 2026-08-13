import dataclasses

@dataclasses.dataclass(frozen=True)
class FwsSpecParameterInfo:
    """
    Summary:
        関数またはメソッドの引数情報を保持するエンティティクラス。
    """
    name:str
    type_name:str
    description:str
