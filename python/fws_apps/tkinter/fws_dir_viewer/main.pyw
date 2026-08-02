"""
Summary:
    ディレクトリビューアーアプリケーションのエントリーポイントモジュール。
ScreenName: ディレクトリビューアー画面
"""

import os
import sys

# パッケージおよび共通ライブラリへのインポート用パス解決
project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), "../../../"))
if project_root not in sys.path:
    sys.path.append(project_root)

import fws_apps.tkinter.fws_dir_viewer.views.events.fws_app_events as fws_app_events


def main():
    """
    Summary:
        アプリケーションのイベントハンドラークラスのインスタンスを生成して起動します。
    """
    app = fws_app_events.FwsAppEvents()
    app.view.mainloop()


if __name__ == "__main__":
    main()
