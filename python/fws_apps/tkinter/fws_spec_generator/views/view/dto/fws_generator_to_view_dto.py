"""不要想定"""
"""
Summary:
    モジュールの概要を記述する。
ScreenName: 画面の日本語物理名
Attachment: 添付設計書のパス
"""

from dataclasses import dataclass

@dataclass(frozen=True)
class FwsGeneratorToViewDTO:
    """Logic から画面への表示用 DTO"""

    source_dir: str
    output_dir: str
    log_text: str
    status_message: str
