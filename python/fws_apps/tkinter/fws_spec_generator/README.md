# 仕様書自動生成ツール (fws_spec_generator)

PythonソースコードをAST（抽象構文木）で解析し、Javadoc風のHTML仕様書を自動生成するデスクトップツールです。
fws_todo_app の設計パターン（DTO、Entityの分離および logics フォルダ構成）に合わせて再編成されています。

## 機能概要

- ソースコード解析: 指定されたフォルダ配下のPythonファイルを再帰的に探索し、クラス、属性、メソッド、関数、docstringをASTで解析します。
- 仕様書HTML生成: 解析されたデータを元に、サイドバーナビゲーション付きの使いやすい Javadoc 風 HTML ファイルを自動生成します。
- UI機能: 暗めのテーマによる画面構成、フォルダ選択用ダイアログ、リアルタイムの実行ログ表示、ステータス表示が備わっています。

## ディレクトリ構造

- main.py: アプリの起動エントリーポイント
- contents/fws_contents.py: HTML/CSS の静的テンプレート定数
- logics/: ビジネスロジックレイヤー
  - fws_app_logic.py: ロジックのメイン中継および config.json の保存管理
  - fws_spec_parser_service.py: ソースコード解析サービス
  - fws_spec_generator_service.py: 仕様書HTML生成サービス
- models/: データモデルレイヤー
  - entity/fws_spec_entities.py: 解析データのドメインモデル
  - entity/fws_generator_entity.py: UI状態管理用エンティティ
  - dto/fws_generator_from_view_dto.py: 画面からLogicへの入力用DTO
  - dto/fws_generator_to_view_dto.py: Logicから画面への表示用DTO
- views/: UI表示レイヤー
  - view/fws_app_view.py: ウィジェット配置およびGetter/Setter
  - event/fws_app_events.py: 画面の操作イベントハンドラー
- tests/fws_test_app.py: テストコード

## 起動方法

以下のコマンドを実行してアプリケーションを起動します。

python fws_apps/tkinter/fws_spec_generator/main.py
（※.venv環境がある場合は .venv/Scripts/python を使用してください）

## テスト実行方法

以下のコマンドで単体テストを実行できます。

python -m unittest fws_apps/tkinter/fws_spec_generator/tests/fws_test_app.py
