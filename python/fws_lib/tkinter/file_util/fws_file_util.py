"""解析中"""
"""
Summary:
    ファイルおよびディレクトリ操作を管理する共通ユーティリティモジュール。
ScreenName: 共通ユーティリティ
"""
import pathlib
from typing import Tuple

def create_directory(target_path: str)->None:
    """
    Summary:
        指定されたパスにディレクトリを作成します。親ディレクトリが存在しない場合も一括で作成します。
    Description:
        pathlib.Path.mkdirを利用し、parents=Trueおよびexist_ok=Trueを指定することで、
        呼び出し元での事前の存在チェックや階層ごとの作成処理を不要にします。
    Args:
        target_path: str - 作成対象のディレクトリの絶対パスまたは相対パス
    Returns:
        None - ディレクトリ作成が成功したとき
    UserAction:
        ファイル保存等、対象ディレクトリへのアクセスが必要な処理が実行された時 - 対象のディレクトリを親要素含めて安全に作成する。
    """
    path_obj = pathlib.Path(target_path)
    path_obj.mkdir(parents=True, exist_ok=True)

def is_directory_exists(target_path: str) -> bool:
    """
    Summary:
        指定されたパスにディレクトリが存在するかどうかを判定します。
    Description:
        パスが存在し、かつディレクトリである場合のみTrueを返します。
        ファイルが同名で存在している場合や、権限エラー等でアクセスできない場合はFalseを返します。
    Args:
        target_path: str - 存在チェックを行う対象の絶対パスまたは相対パス
    Returns:
        bool - ディレクトリが存在する場合はTrue、存在しない（またはファイルである、アクセス不可の）場合はFalse。
    UserAction:
        ディレクトリへのアクセス前処理時 - 指定パスのディレクトリの有無を確認し、後続処理の分岐に利用する。
    """
    path_obj = pathlib.Path(target_path)
    return path_obj.is_dir()
        