"""解析中"""
"""
Summary:
    Tkinterのメッセージボックス関連の共通ユーティリティモジュール。
"""
from tkinter import messagebox

def show_info(title: str, message: str, parent=None) -> None:
    """
    Summary:
        情報メッセージダイアログを表示します。
    Args:
        title: str - ダイアログのタイトル
        message: str - 表示するメッセージ内容
        parent: Tk/Toplevel - 親ウィンドウ（指定すると親の中央に表示される）
    """
    messagebox.showinfo(title, message, parent=parent)

def show_warning(title: str, message: str, parent=None) -> None:
    """
    Summary:
        警告メッセージダイアログを表示します。
    Args:
        title: str - ダイアログのタイトル
        message: str - 表示するメッセージ内容
        parent: Tk/Toplevel - 親ウィンドウ
    """
    messagebox.showwarning(title, message, parent=parent)

def show_error(title: str, message: str, parent=None) -> None:
    """
    Summary:
        エラーメッセージダイアログを表示します。
    Args:
        title: str - ダイアログのタイトル
        message: str - 表示するメッセージ内容
        parent: Tk/Toplevel - 親ウィンドウ
    """
    messagebox.showerror(title, message, parent=parent)

def ask_yes_no(title: str, message: str, parent=None) -> bool:
    """
    Summary:
        「はい / いいえ」の確認ダイアログを表示します。
    Args:
        title: str - ダイアログのタイトル
        message: str - 表示する確認メッセージ内容
        parent: Tk/Toplevel - 親ウィンドウ
    Returns:
        bool - 「はい」が選択された場合はTrue、「いいえ」の場合はFalse
    """
    return messagebox.askyesno(title, message, parent=parent)

def ask_ok_cancel(title: str, message: str, parent=None) -> bool:
    """
    Summary:
        「OK / キャンセル」の確認ダイアログを表示します。
    Args:
        title: str - ダイアログのタイトル
        message: str - 表示する確認メッセージ内容
        parent: Tk/Toplevel - 親ウィンドウ
    Returns:
        bool - 「OK」が選択された場合はTrue、「キャンセル」の場合はFalse
    """
    return messagebox.askokcancel(title, message, parent=parent)