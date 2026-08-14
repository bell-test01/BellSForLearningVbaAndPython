"""解析中"""
from dataclasses import asdict, dataclass
import uuid


from fws_apps.tkinter.fws_spec_generator.views.view.dto import fws_generator_from_view_dto
from fws_apps.tkinter.fws_spec_generator.views.view.dto import fws_generator_to_view_dto


@dataclass
class FwsGeneratorEntity:
    """画面状態管理用 Entity"""

    id: str
    source_dir: str
    output_dir: str
    log_text: str = ""
    status_message: str = "準備完了"

    def __post_init__(self):
        if not self.id:
            self.id = str(uuid.uuid4())

    @classmethod
    def from_view_dto(
        cls, dto: fws_generator_from_view_dto.FwsGeneratorFromViewDTO
    ) -> "FwsGeneratorEntity":
        """FwsGeneratorFromViewDTO から Entity を生成"""
        return cls(
            id="",
            source_dir=dto.source_dir.strip(),
            output_dir=dto.output_dir.strip(),
        )

    def to_view_dto(self) -> fws_generator_to_view_dto.FwsGeneratorToViewDTO:
        """Entity から FwsGeneratorToViewDTO へ変換"""
        return fws_generator_to_view_dto.FwsGeneratorToViewDTO(
            source_dir=self.source_dir,
            output_dir=self.output_dir,
            log_text=self.log_text,
            status_message=self.status_message,
        )

    def to_dict(self) -> dict:
        """JSON 保存用辞書変換"""
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict) -> "FwsGeneratorEntity":
        """JSON 辞書からの復元"""
        return cls(
            id=data.get("id", ""),
            source_dir=data.get("source_dir", ""),
            output_dir=data.get("output_dir", ""),
            log_text=data.get("log_text", ""),
            status_message=data.get("status_message", "準備完了"),
        )
