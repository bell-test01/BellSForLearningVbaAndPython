"""解析中"""
"""
Summary:
    Tkinterのファイルダイアログ関連の共通ユーティリティモジュール。
"""
import os
import tkinter.filedialog as filedialog

def ask_directory(title: str, initialdir: str = None) -> str:
    """
    Summary:
        フォルダ選択ダイアログを表示し、選択したフォルダのパスを返します。
    Args:
        title: str - ダイアログのタイトル
        initialdir: str - 初期ディレクトリパス（未指定の場合はデフォルト）
    Returns:
        str - 選択したフォルダのパス（未選択の場合は空文字）
    """
    selected:str = filedialog.askdirectory(title=title,initialdir=initialdir)
    
    result:str = ""
    if selected:
        result = os.path.normpath(selected)
        
    return result

def ask_open_filename(title: str, initialdir: str = None, filetypes: list = None) -> str:
    """
    Summary:
        ファイル選択ダイアログ（単一選択）を表示し、選択したファイルのパスを返します。
    Args:
        title: str - ダイアログのタイトル
        initialdir: str - 初期ディレクトリパス（未指定の場合はデフォルト）
        filetypes: list - 選択可能なファイル種別のリスト (例: [("Text files", "*.txt"), ("All files", "*.*")])
    Returns:
        str - 選択したファイルのパス（未選択の場合は空文字）
    """
    if filetypes is None:
        filetypes = [("All files", "*.*")]
    
    selected = filedialog.askopenfilename(title=title, initialdir=initialdir, filetypes=filetypes)
    if selected:
        return os.path.normpath(selected)
    return ""

def ask_open_filenames(title: str, initialdir: str = None, filetypes: list = None) -> list[str]:
    """
    Summary:
        ファイル選択ダイアログ（複数選択）を表示し、選択したファイルのパスのリストを返します。
    Args:
        title: str - ダイアログのタイトル
        initialdir: str - 初期ディレクトリパス（未指定の場合はデフォルト）
        filetypes: list - 選択可能なファイル種別のリスト
    Returns:
        list[str] - 選択したファイルのパスのリスト（未選択の場合は空リスト）
    """
    if filetypes is None:
        filetypes = [("All files", "*.*")]
        
    selected = filedialog.askopenfilenames(title=title, initialdir=initialdir, filetypes=filetypes)
    if selected:
        return [os.path.normpath(p) for p in selected]
    return []

def ask_save_as_filename(title: str, initialdir: str = None, defaultextension: str = None, filetypes: list = None, initialfile: str = None) -> str:
    """
    Summary:
        ファイル保存ダイアログを表示し、指定した保存先のファイルパスを返します。
    Args:
        title: str - ダイアログのタイトル
        initialdir: str - 初期ディレクトリパス（未指定の場合はデフォルト）
        defaultextension: str - デフォルトの拡張子（例: ".txt"）
        filetypes: list - 選択可能なファイル種別のリスト
        initialfile: str - 初期ファイル名
    Returns:
        str - 選択した保存先のファイルパス（未選択の場合は空文字）
    """
    if filetypes is None:
        filetypes = [("All files", "*.*")]
        
    selected = filedialog.asksaveasfilename(
        title=title,
        initialdir=initialdir,
        defaultextension=defaultextension,
        filetypes=filetypes,
        initialfile=initialfile
    )
    if selected:
        return os.path.normpath(selected)
    return ""