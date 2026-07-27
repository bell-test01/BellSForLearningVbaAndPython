from dataclasses import dataclass


@dataclass(frozen=True)
class FwsGeneratorToViewDTO:
    """Logic から画面への表示用 DTO"""

    source_dir: str
    output_dir: str
    log_text: str
    status_message: str
