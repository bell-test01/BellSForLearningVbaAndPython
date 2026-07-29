"""
Summary:
    SQLite Viewerアプリケーションのエントリーポイントモジュール。
ScreenName: SQLite Viewerアプリケーション
"""

import os
import sys

# パッケージおよび共通ライブラリへのインポート用パス解決
project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), "../../../"))
if project_root not in sys.path:
    sys.path.append(project_root)

#app本体をインポート
from fws_apps.tkinter.fws_sqlite_viewer.views.event import fws_app_events

def main():
    """
    Summary:
        アプリケーションのイベントハンドラークラスのインスタンスを生成して起動します。
    """
    fws_app_events_obj = fws_app_events.FwsAppEvents()
    fws_app_events_obj.view.mainloop()

if __name__ == "__main__":
    main()
