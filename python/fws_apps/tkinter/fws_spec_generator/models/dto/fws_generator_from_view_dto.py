from dataclasses import dataclass


@dataclass(frozen=True)
class FwsGeneratorFromViewDTO:
    """画面から Logic への入力用 DTO"""

    source_dir: str
    output_dir: str
