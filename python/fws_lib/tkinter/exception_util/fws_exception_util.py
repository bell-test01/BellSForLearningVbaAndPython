"""解析中"""
from fws_lib.tkinter.msg_util import fws_msg_util

def test_except(e:Exception)->None:
    """
    Summary:
        例外エラー発生時の共通処理を実装したい
    Args:
        e: Exception - 例外オブジェクト
    """
    fws_msg_util.show_error("エラー", str(e))