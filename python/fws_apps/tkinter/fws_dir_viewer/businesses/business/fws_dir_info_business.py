"""
Summary:
"""

import os
from datetime import datetime

import fws_apps.tkinter.fws_dir_viewer.businesses.entity.fws_dir_info_entity as fws_dir_info_entity


class FwsDirInfoBusiness:
    """
    Summary:
        指定されたパスのディレクトリ配下のアイテム情報を取得するクラス。
    """

    def get_dir_info(self, entity: fws_dir_info_entity.FwsDirInfoEntity) -> fws_dir_info_entity.FwsDirInfoEntity:
        """
        Summary:
            エンティティに設定されたパスを元に、ディレクトリ配下のファイルやディレクトリ一覧を取得し、結果をエンティティに設定して返します。
        Args:
            entity: FwsDirInfoEntity - 処理対象のパスが設定されたエンティティ。
        Returns:
            FwsDirInfoEntity - 結果が格納されたエンティティ。
        """
        path = entity.dir_path.strip()

        if not path:
            entity.is_success = False
            entity.message = "エラー: ディレクトリパスを入力してください。"
            return entity

        if not os.path.exists(path):
            entity.is_success = False
            entity.message = f"エラー: 指定されたパスが存在しません。({path})"
            return entity

        if not os.path.isdir(path):
            entity.is_success = False
            entity.message = f"エラー: 指定されたパスはディレクトリではありません。({path})"
            return entity

        items = []
        try:
            for item in os.listdir(path):
                full_path = os.path.join(path, item)
                is_dir = os.path.isdir(full_path)
                item_type = "Dir" if is_dir else "File"
                
                try:
                    stat = os.stat(full_path)
                    size = "-" if is_dir else str(stat.st_size)
                    mtime = datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
                except Exception:
                    size = "不明"
                    mtime = "不明"

                items.append((item, item_type, size, mtime, full_path))

            # ディレクトリを先に、次にファイルをソート
            items.sort(key=lambda x: (x[1] == "File", x[0].lower()))
            
            entity.items = items
            entity.is_success = True
            entity.message = f"成功: {len(items)} 件のアイテムを取得しました。"

        except Exception as e:
            entity.is_success = False
            entity.message = f"エラー: 情報の取得中に例外が発生しました。({str(e)})"

        return entity