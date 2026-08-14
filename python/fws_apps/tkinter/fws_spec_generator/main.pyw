"""解析済"""
"""
Summary:
    仕様書自動生成ツールのインスタンスを生成して起動する。
"""
import pathlib
import sys

# パッケージおよび共通ライブラリへのインポート用パス解決
project_root = str(pathlib.Path(__file__).resolve().parents[3])
if project_root not in sys.path:
    sys.path.append(project_root)

from fws_apps.tkinter.fws_spec_generator.views.event import fws_app_event

def main()->None:
    """
    Summary:
        アプリ起動
    Description:
        仕様書自動生成ツールを起動する。
    Returns:
        None - 無し
    """
    app = fws_app_event.FwsAppEvent()
    app.fws_app_view_obj.mainloop()

if __name__ == "__main__":
    main()
