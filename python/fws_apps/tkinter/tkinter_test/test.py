import tkinter as tk
from tkinter import ttk
from tkinter import messagebox, filedialog, colorchooser

# ==============================================================================
# 画面への配置方法
# ==============================================================================
"""
1. grid (グリッド・格子状配置) ※実務で一番よく使う推奨配置
├── row / column         : 【影響】配置するマスの「行番号」と「列番号」を指定 (0始まり)
│                          └─ 例: grid(row=0, column=1)
├── rowspan / columnspan : 【影響】縦方向・横方向に何マス分結合して繋げるかを指定
│                          └─ 例: grid(row=0, column=0, columnspan=2)  # 2列分を結合
├── sticky               : 【影響】割り当てられたマス目の中で「どの方向（壁）に寄せる・伸ばす」かを指定
│                          │      n(上), s(下), e(右), w(左) の組み合わせ
│                          ├─ 単体寄せ: sticky="w" (左寄せ), sticky="e" (右寄せ)
│                          ├─ 横伸長  : sticky="ew" (マスの左右いっぱいに引き伸ばす)
│                          └─ 全方位  : sticky="nsew" (マス全体の上下左右いっぱいに拡大)
├── padx / pady          : 【影響】パーツの「外側（枠の外）」に設ける余白（すき間）をピクセル指定
│                          ├─ 均等余白: padx=10, pady=5  # 横に10px, 縦に5pxの余白
│                          └─ 個別余白: padx=(10, 20)     # (左に10px, 右に20px) の非対称余白
└── ipadx / ipady        : 【影響】パーツの「内側（枠の中）」に設ける余白（内側から膨らませる）を指定
                           └─ 例: ipadx=5, ipady=10

2. pack (パック・流し込み配置) ※シンプルな縦並べ・横並べに便利
├── side                 : 【影響】親枠の「どの端（上下左右）」から順番に詰めて並べるかを指定
│                          └─ 例: pack(side="top")  # "top"(上から下), "bottom"(下から上), "left"(左から右), "right"(右から左)
├── fill                 : 【影響】余ったスペースに対して「どちらの方向に引っ張って広げるか」を指定
│                          └─ 例: pack(fill="x")  # "none"(伸縮なし), "x"(横一杯), "y"(縦一杯), "both"(縦横両方)
├── expand               : 【影響】親ウィンドウがリサイズされた際、生まれた「余白エリアを配分して押し広げるか」を設定
│                          └─ 例: pack(expand=True)  # True: 余白を埋めるように広がる, False: 自身のサイズを維持
├── anchor               : 【影響】割り当てられた領域内で「どの位置（方角）に固定して寄せるか」を指定
│                          └─ 例: anchor="nw"  # "n"(北), "s"(南), "e"(東), "w"(西), "nw"(左上), "se"(右下), "center"(中央)
├── padx / pady          : 【影響】パーツの「外側（枠の外）」に設ける余白（すき間）を指定
│                          └─ 例: pack(padx=10, pady=5)  /  pack(padx=(10, 0)) # 左だけ10px
└── ipadx / ipady        : 【影響】パーツの「内側（枠の中）」に設ける余白を指定
                           └─ 例: pack(ipadx=10, ipady=5)

3. (補足) 親要素側のグリッド伸縮設定 (columnconfigure / rowconfigure)
   ※ grid で sticky="ew" などを使っても広がり切らない場合、親枠側に設定します。
├── columnconfigure(...) : 【影響】特定「列」の伸縮比率（重み）を設定
│                          └─ 例: root.columnconfigure(1, weight=1)  # 1列目がウィンドウサイズに応じて伸び縮みする
└── rowconfigure(...)    : 【影響】特定「行」の伸縮比率（重み）を設定
                           └─ 例: root.rowconfigure(0, weight=1)     # 0行目が縦方向に伸び縮みする
"""

# ==============================================================================
# 1. アプリ最上位・ウインドウ（土台）
# ==============================================================================
"""
1. アプリ最上位・ウインドウ（土台）
├── [1] tk.Tk / [2] tk.Toplevel (ウインドウ設定)
│   ├── title(...)        : 【影響】タイトルバーの文字列を変更
│   │                       └─ 例: title("マイアプリ")
│   ├── geometry(...)     : 【影響】初期の幅・高さと表示画面位置を設定
│   │                       └─ 例: geometry("800x600+100+100")  # 幅x高さ+X位置+Y位置
│   ├── resizable(...)    : 【影響】ユーザーによるウインドウ枠の拡大縮小を制限
│   │                       └─ 例: resizable(True, False)      # 横可変, 縦固定
│   ├── minsize / maxsize : 【影響】ウインドウサイズ変更時の最小/最大制限を設定
│   │                       └─ 例: minsize(400, 300) / maxsize(1200, 900)
│   ├── config(...)       : 【影響】ウインドウ全体の背景色を変更
│   │                       └─ 例: config(bg="#f0f0f0")
│   ├── attributes(...)   : 【影響】OS固有の特殊効果（透明度・最前面化等）を適用
│   │                       └─ 例: attributes("-alpha", 0.95, "-topmost", True)
│   └── protocol(...)     : 【影響】「×ボタン」で閉じる際の割り込み処理を設定
│                           └─ 例: protocol("WM_DELETE_WINDOW", on_close_func)
"""
# --- [1] tk.Tk ---
root = tk.Tk()
root.title("Tkinter 全アイテム実コード一覧")
root.geometry("600x800")

# --- [2] tk.Toplevel ---
# 別ウインドウ（サブ画面）を開く
sub_window = tk.Toplevel(root)
sub_window.title("サブウインドウ")
sub_window.geometry("300x150")


# ==============================================================================
# 2. 構造・レイアウト用コンテナ（他の要素を載せる親）
# ==============================================================================
"""
2. 構造・レイアウト用コンテナ（枠・分割）
├── [3] ttk.Frame (透明パネル)
│   ├── padding           : 【影響】枠の内側に設けるすき間（余白）を指定
│   │                       └─ 例: padding=10  /  padding=(10, 5, 10, 5) # 左,上,右,下
│   ├── width / height    : 【影響】パネルの幅・高さをピクセル単位で明示的に固定
│   │                       └─ 例: width=300, height=200
│   ├── relief            : 【影響】パネル外枠の立体感スタイルを変更
│   │                       └─ 例: relief="sunken"  # "flat", "sunken", "raised", "groove", "ridge"
│   └── style             : 【影響】定義済みのttkデザインスタイルを適用
│                           └─ 例: style="Custom.TFrame"
"""
# --- [3] ttk.Frame ---
# 透明なベースパネル
frame = ttk.Frame(root, padding=10)
frame.pack(fill="x")

"""
├── [4] ttk.LabelFrame (タイトル枠付きパネル)
│   ├── text              : 【影響】枠線の上の見出しタイトル文字を設定
│   │                       └─ 例: text="ユーザー情報"
│   ├── labelanchor       : 【影響】見出しタイトルの配置場所を指定
│   │                       └─ 例: labelanchor="nw"  # "nw"(左上), "n"(中央上), "ne"(右上) 等
│   ├── padding           : 【影響】枠内部のコンテンツとの余白を指定
│   │                       └─ 例: padding=10
│   └── relief / style    : 【影響】枠線の見た目立体感スタイルを変更
│                           └─ 例: relief="groove", style="Group.TLabelframe"
"""
# --- [4] ttk.LabelFrame ---
# タイトル枠付きパネル
labelframe = ttk.LabelFrame(root, text="グループ枠", padding=10)
labelframe.pack(fill="x", padx=10, pady=5)

"""
├── [5] ttk.Notebook (タブ切り替えパネル)
│   ├── padding           : 【影響】タブ枠全体の外側余白を指定
│   │                       └─ 例: padding=5
│   ├── add(...)          : 【影響】新しいタブページを追加（タイトルやアイコン付与）
│   │                       └─ 例: add(tab_frame, text="設定", image=icon_img, compound="left")
│   ├── tab(...)          : 【影響】特定タブの表示・非表示・無効化状態を変更
│   │                       └─ 例: tab(tab_frame, state="disabled")  # "normal", "disabled", "hidden"
│   └── enable_traversal(): 【影響】Ctrl+Tabキー等でのタブ移動を有効化
│                           └─ 例: enable_traversal()
"""
# --- [5] ttk.Notebook ---
# タブ切り替えパネル
notebook = ttk.Notebook(root)
tab1 = ttk.Frame(notebook)
tab2 = ttk.Frame(notebook)
notebook.add(tab1, text="タブ1")
notebook.add(tab2, text="タブ2")
notebook.pack(fill="x", padx=10, pady=5)

"""
└── [6] ttk.PanedWindow (ドラッグ可変分割)
    ├── orient            : 【影響】分割境界線の方向（左右分割か上下分割か）を設定
    │                       └─ 例: orient=tk.HORIZONTAL  # tk.HORIZONTAL(左右) / tk.VERTICAL(上下)
    └── add(...)          : 【影響】子パネルを追加（リサイズ時の伸縮優先度も指定）
                            └─ 例: add(left_frame, weight=1)  # weight: ウィンドウ変更時の伸縮比率
"""
# --- [6] ttk.PanedWindow ---
# マウスドラッグで可変な分割パネル
paned = ttk.PanedWindow(root, orient=tk.HORIZONTAL)
left_pane = ttk.Frame(paned, width=100, height=50)
right_pane = ttk.Frame(paned, width=100, height=50)
paned.add(left_pane)
paned.add(right_pane)
paned.pack(fill="x", padx=10, pady=5)


# ==============================================================================
# 3. メニュー構造
# ==============================================================================

"""
├── [7] tk.Menu (メニューバー / ドロップダウン)
│   ├── tearoff           : 【影響】メニュー上部の切り離し点線（ウインドウ化機能）の有無
│   │                       └─ 例: tearoff=0  # 0: なし(推奨), 1: あり
│   ├── bg / fg           : 【影響】メニュー項目の背景色と文字色を変更
│   │                       └─ 例: bg="#333333", fg="#ffffff"
│   ├── add_command(...)  : 【影響】通常クリック項目（ラベル・処理・ショートカット表示）を追加
│   │                       └─ 例: add_command(label="開く", command=open_func, accelerator="Ctrl+O")
│   ├── add_checkbutton() : 【影響】ON/OFFチェックがつくメニュー項目を追加
│   │                       └─ 例: add_checkbutton(label="グリッド表示", variable=grid_var)
│   ├── add_radiobutton() : 【影響】1つだけ選択できるラジオメニュー項目を追加
│   │                       └─ 例: add_radiobutton(label="モードA", variable=mode_var, value="A")
│   └── add_cascade(...)  : 【影響】階層型のサブメニュー（ドロップダウン）を接続
│                           └─ 例: add_cascade(label="ファイル", menu=file_menu)
"""
# --- [7] tk.Menu ---
# 最上部メニューバー
menubar = tk.Menu(root)
file_menu = tk.Menu(menubar, tearoff=0)
file_menu.add_command(label="開く", command=lambda: print("メニュークリック"))
file_menu.add_separator()
file_menu.add_command(label="終了", command=root.quit)
menubar.add_cascade(label="ファイル", menu=file_menu)
root.config(menu=menubar)

"""
└── [8] ttk.Menubutton (画面内メニューボタン)
    ├── text / image      : 【影響】ボタンの上に表示するテキストやアイコンを設定
    │                       └─ 例: text="アクション", image=icon_img
    ├── menu              : 【影響】クリック時に開く `tk.Menu` オブジェクトを紐付け
    │                       └─ 例: menu=sub_menu_object
    └── direction         : 【影響】メニューが開く方向（下・上・左・右）を指定
                            └─ 例: direction="below"  # "below", "above", "left", "right"
"""
# --- [8] ttk.Menubutton ---
# 画面内に配置するドロップダウンボタン
menubtn = ttk.Menubutton(labelframe, text="操作メニュー")
mb_menu = tk.Menu(menubtn, tearoff=0)
mb_menu.add_command(label="処理A", command=lambda: print("A実行"))
menubtn["menu"] = mb_menu
menubtn.pack(anchor="w")


# ==============================================================================
# 4. 主要入力・操作系ウィジェット
# ==============================================================================

"""
├── [9] ttk.Label (文字・画像表示)
│   ├── text              : 【影響】表示する固定文字列を設定
│   │                       └─ 例: text="ステータス: 正常"
│   ├── textvariable      : 【影響】動的に変化する `tk.StringVar` 変数とテキストをリアルタイム連動
│   │                       └─ 例: textvariable=status_var
│   ├── image             : 【影響】文字の代わりに（または一緒に）画像を表示
│   │                       └─ 例: image=photo_img
│   ├── compound          : 【影響】画像と文字列を併記する際の位置関係を指定
│   │                       └─ 例: compound="left"  # "top", "bottom", "left", "right", "center"
│   ├── anchor            : 【影響】割り当てられた枠内でのテキスト寄せ方向を指定
│   │                       └─ 例: anchor="w"     # "n"(上), "s"(下), "e"(右), "w"(左), "center"
│   ├── justify           : 【影響】複数行テキストの改行揃え方向を指定
│   │                       └─ 例: justify="left" # "left", "center", "right"
│   ├── font              : 【影響】フォントの種類・サイズ・装飾を設定
│   │                       └─ 例: font=("Meiryo", 11, "bold")
│   └── wraplength        : 【影響】指定ピクセル幅を超えた場合に自動折り返し
│                           └─ 例: wraplength=200
"""
# --- [9] ttk.Label ---
lbl = ttk.Label(labelframe, text="ラベル初期値")
lbl.pack(anchor="w")
lbl["text"] = "【Label】変更後のテキスト"  # 変更

"""
├── [10] ttk.Entry (1行入力)
│   ├── textvariable      : 【影響】入力テキストとリアルタイム同期する `tk.StringVar` を紐付け
│   │                       └─ 例: textvariable=entry_var
│   ├── width             : 【影響】入力枠の横幅を「標準文字数単位」で指定
│   │                       └─ 例: width=30
│   ├── show              : 【影響】入力文字を別の記号に置き換えて表示（パスワード用）
│   │                       └─ 例: show="*"
│   ├── state             : 【影響】入力枠の編集可否状態を変更
│   │                       └─ 例: state="normal" # "normal"(通常), "disabled"(禁止), "readonly"(読取専用)
│   ├── justify           : 【影響】カーソル・入力テキストの寄せ方向を指定
│   │                       └─ 例: justify="right"# 右寄せ入力
│   └── validate / valcmd : 【影響】リアルタイムに入力値チェック（数字のみ等）を行うフックを設定
│                           └─ 例: validate="key", validatecommand=(root.register(check_func), "%P")
"""
# --- [10] ttk.Entry ---
entry_var = tk.StringVar(value="初期入力値")
entry = ttk.Entry(labelframe, textvariable=entry_var)
entry.pack(fill="x")
entry_var.trace_add("write", lambda *a: print(f"Entry変更: {entry_var.get()}"))  # イベント
entry.bind("<Return>", lambda e: print("Enterキー確定"))

"""
├── [11] ttk.Button (ボタン)
│   ├── text / image      : 【影響】ボタンの上に表示する文字列や画像・配置を設定
│   │                       └─ 例: text="送信", image=icon_img, compound="left"
│   ├── command           : 【影響】ボタンクリック時に実行する関数を指定
│   │                       └─ 例: command=submit_func
│   ├── default           : 【影響】Enterキー押下でこのボタンを反応させるかどうかのデフォルト設定
│   │                       └─ 例: default="active"  # "active", "normal", "disabled"
│   └── state             : 【影響】ボタンをグレーアウト（押せない状態）にするか設定
│                           └─ 例: state="disabled"
"""
# --- [11] ttk.Button ---
btn = ttk.Button(labelframe, text="実行ボタン", command=lambda: print("ボタンクリック"))
btn.pack(anchor="e")

"""
├── [12] ttk.Checkbutton (チェックボックス)
│   ├── text              : 【影響】チェックボックスの横に表示するラベル文字列
│   │                       └─ 例: text="利用規約に同意する"
│   ├── variable          : 【影響】ON/OFFの状態値を記憶する変数（`tk.BooleanVar`等）を指定
│   │                       └─ 例: variable=agree_var
│   ├── onvalue / offvalue: 【影響】チェックON時/OFF時に変数に設定される値をカスタマイズ
│   │                       └─ 例: onvalue=True, offvalue=False (または onvalue="YES", offvalue="NO")
│   ├── command           : 【影響】切り替わった瞬間に実行する関数を指定
│   │                       └─ 例: command=on_toggle_func
│   └── state             : 【影響】チェックボックス操作の有効/無効を変更
│                           └─ 例: state="normal"
"""
# --- [12] ttk.Checkbutton ---
check_var = tk.BooleanVar(value=True)
chk = ttk.Checkbutton(labelframe, text="チェックボックス", variable=check_var, command=lambda: print(f"Check: {check_var.get()}"))
chk.pack(anchor="w")

"""
├── [13] ttk.Radiobutton (ラジオボタン)
│   ├── text              : 【影響】各ラジオボタンの選択肢ラベル文字列
│   │                       └─ 例: text="男性"
│   ├── variable          : 【影響】グループ内で同じものを指定し、1つだけ選ばれる関係を作る変数
│   │                       └─ 例: variable=gender_var
│   ├── value             : 【影響】このボタンが選ばれたときに `variable` に代入される値
│   │                       └─ 例: value="M"
│   └── command           : 【影響】選択が変更された瞬間に実行する関数を指定
│                           └─ 例: command=on_select_func
"""
# --- [13] ttk.Radiobutton ---
radio_var = tk.StringVar(value="A")
rb1 = ttk.Radiobutton(labelframe, text="選択A", value="A", variable=radio_var, command=lambda: print(f"Radio: {radio_var.get()}"))
rb2 = ttk.Radiobutton(labelframe, text="選択B", value="B", variable=radio_var, command=lambda: print(f"Radio: {radio_var.get()}"))
rb1.pack(anchor="w")
rb2.pack(anchor="w")

"""
└── [14] ttk.Combobox (ドロップダウン入力)
    ├── textvariable      : 【影響】選択・入力された値が代入される `tk.StringVar`
    │                       └─ 例: textvariable=combo_var
    ├── values            : 【影響】ドロップダウンに表示する選択肢の一覧リストを設定
    │                       └─ 例: values=["東京", "大阪", "名古屋"]
    ├── state             : 【影響】「手入力可」か「選択のみ（直入力不可）」かを変更
    │                       └─ 例: state="readonly"  # "normal":手入力可, "readonly":選択のみ
    └── height            : 【影響】ドロップダウンを開いた際に一画面で表示する最大行数
                            └─ 例: height=5
"""
# --- [14] ttk.Combobox ---
combo_var = tk.StringVar(value="東京")
combo = ttk.Combobox(labelframe, textvariable=combo_var, values=["東京", "大阪", "名古屋"])
combo.pack(fill="x")
combo.bind("<<ComboboxSelected>>", lambda e: print(f"Combo選択: {combo_var.get()}"))  # イベント


# ==============================================================================
# 5. 数値・範囲設定ウィジェット
# ==============================================================================

"""
├── [15] ttk.Spinbox (数値上下ボックス)
│   ├── from_ / to        : 【影響】選択可能な数値の最小値と最大値を設定
│   │                       └─ 例: from_=1, to=100
│   ├── increment         : 【影響】矢印ボタンを押したときに増減するステップ幅
│   │                       └─ 例: increment=5
│   ├── textvariable      : 【影響】現在の設定値を管理する数値変数（`tk.IntVar`等）
│   │                       └─ 例: textvariable=num_var
│   ├── values            : 【影響】数値ではなく固定の文字列リスト（"小","中","大"）から選ばせる場合に使用
│   │                       └─ 例: values=("小", "中", "大")
│   └── wrap              : 【影響】最大値（または最小値）に達したときに反対側にループさせるか
│                           └─ 例: wrap=True
"""
# --- [15] ttk.Spinbox ---
spin_var = tk.IntVar(value=5)
spin = ttk.Spinbox(labelframe, from_=0, to=10, textvariable=spin_var)
spin.pack(anchor="w")
spin_var.trace_add("write", lambda *a: print(f"Spinbox: {spin_var.get()}"))

"""
├── [16] ttk.Scale (スライダー)
│   ├── from_ / to        : 【影響】スライダーの両端の範囲数値を指定
│   │                       └─ 例: from_=0.0, to=1.0
│   ├── orient            : 【影響】スライダーの配置方向（横向きか縦向きか）
│   │                       └─ 例: orient="horizontal" # "horizontal", "vertical"
│   ├── variable          : 【影響】スライダーの位置数値を連動させる `tk.DoubleVar`
│   │                       └─ 例: variable=scale_var
│   ├── command           : 【影響】つまみをドラッグ中にリアルタイム実行される関数（引数に数値が入る）
│   │                       └─ 例: command=lambda val: print(f"現在値: {float(val):.2f}")
│   └── value             : 【影響】プロパティ経由で直接数値を割り当て・読み出し
│                           └─ 例: scale.value = 0.5
"""
# --- [16] ttk.Scale ---
scale_var = tk.DoubleVar(value=50.0)
scale = ttk.Scale(labelframe, from_=0, to=100, variable=scale_var, command=lambda val: print(f"Scale: {float(val):.1f}"))
scale.pack(fill="x")

"""
└── [17] ttk.Progressbar (プログレスバー)
    ├── orient            : 【影響】バーの進行方向（横伸びか縦伸びか）
    │                       └─ 例: orient="horizontal"
    ├── length            : 【影響】バー全体の表示ピクセル長
    │                       └─ 例: length=300
    ├── mode              : 【影響】進捗指定（確定型）か、左右往復のアニメーション（不確定型）か切り替え
    │                       └─ 例: mode="determinate" # "determinate":進捗量指定, "indeterminate":往復
    ├── maximum           : 【影響】100%完了を意味する最大数値を設定
    │                       └─ 例: maximum=100.0
    └── variable          : 【影響】進捗数値を制御・連動させる `tk.DoubleVar`
                            └─ 例: variable=progress_var
"""
# --- [17] ttk.Progressbar ---
prog_var = tk.DoubleVar(value=70.0)
prog = ttk.Progressbar(labelframe, variable=prog_var, maximum=100)
prog.pack(fill="x")


# ==============================================================================
# 6. 複数行・データ・描画系ウィジェット
# ==============================================================================

"""
├── [18] tk.Text (複数行テキスト編集)
│   ├── width / height    : 【影響】枠の基本サイズを「文字数」と「行数」単位で指定
│   │                       └─ 例: width=40, height=10
│   ├── wrap              : 【影響】枠の右端に達した際の折り返し単位を指定
│   │                       └─ 例: wrap="word"  # "char":文字単位, "word":単語単位, "none":折り返しなし
│   ├── font              : 【影響】エディタ部分のフォントの種類やサイズ
│   │                       └─ 例: font=("Consolas", 10)
│   ├── bg / fg           : 【影響】エディタエリアの背景色と文字色を指定
│   │                       └─ 例: bg="#1e1e1e", fg="#d4d4d4" # ダークモード風
│   ├── insertbackground  : 【影響】入力位置を示す点滅カーソル（キャレット）の色を変更
│   │                       └─ 例: insertbackground="#ffffff"
│   ├── undo              : 【影響】Ctrl+Z / Ctrl+Y による元に戻す・やり直し機能を有効化
│   │                       └─ 例: undo=True
│   └── yscrollcommand    : 【影響】縦スクロールバーの位置とスクロール量を連動
│                           └─ 例: yscrollcommand=scrollbar.set
"""
# --- [18] tk.Text ---
txt = tk.Text(root, height=3)
txt.pack(fill="x", padx=10, pady=5)
txt.insert("1.0", "複数行のテキスト入力エリア\n2行目")  # セット
text_val = txt.get("1.0", tk.END).strip()            # 取得
txt.bind("<KeyRelease>", lambda e: print("Textキー入力"))  # イベント

"""
├── [19] tk.Listbox (リスト一覧選択)
│   ├── selectmode        : 【影響】項目の選択スタイル（単一選択、Ctrlキーでの複数選択など）を変更
│   │                       └─ 例: selectmode="extended"# "single", "browse", "multiple", "extended"
│   ├── height            : 【影響】一度に表示する行数を指定
│   │                       └─ 例: height=6
│   ├── listvariable      : 【影響】リスト全体のデータ（タプル形式）をまとめて連動・差し替え
│   │                       └─ 例: listvariable=items_var # tk.StringVar(value=("A", "B", "C"))
│   ├── selectbackground  : 【影響】選択された行のハイライト背景色を変更
│   │                       └─ 例: selectbackground="#007acc"
│   └── yscrollcommand    : 【影響】スクロールバーとの連動設定
│                           └─ 例: yscrollcommand=scrollbar.set
"""
# --- [19] tk.Listbox ---
lb = tk.Listbox(root, height=3)
lb.pack(fill="x", padx=10, pady=5)
for item in ["項目1", "項目2", "項目3",1,2,3]:
    lb.insert(tk.END, item)
lb.bind("<<ListboxSelect>>", lambda e: print(f"Listbox選択: {lb.get(lb.curselection())}" if lb.curselection() else ""))

"""
├── [20] ttk.Treeview (表・ツリー表示)
│   ├── columns           : 【影響】表の列ID（ヘッダー名）の定義リスト
│   │                       └─ 例: columns=("id", "name", "price")
│   ├── show              : 【影響】階層ツリー列を表示するか、表形式ヘッダーのみを表示するか指定
│   │                       └─ 例: show="headings" # "headings":表のみ, "tree headings":ツリー付き表
│   ├── selectmode        : 【影響】行選択の制限（単一行のみ選択、複数行選択など）
│   │                       └─ 例: selectmode="browse" # "extended", "browse", "none"
│   ├── height            : 【影響】表の高さ（表示行数）を指定
│   │                       └─ 例: height=8
│   └── yscrollcommand / xscrollcommand : 【影響】縦・横の各スクロールバーと連携設定
│                           └─ 例: yscrollcommand=scrollbar.set
"""
# --- [20] ttk.Treeview ---
tree = ttk.Treeview(root, columns=("ID", "Name"), show="headings", height=3)
tree.heading("ID", text="ID",anchor="w")
tree.heading("Name", text="名前")
tree.insert("", tk.END, values=("1", "山田"))
tree.insert("", tk.END, values=("2", "佐藤"))
tree.insert("", tk.END, values=("3", "佐藤"))
tree.insert("", tk.END, values=("4", "佐藤"))
tree.pack(fill="x", padx=10, pady=5)
tree.bind("<<TreeviewSelect>>", lambda e: print(f"Treeview選択: {tree.item(tree.selection())['values']}" if tree.selection() else ""))

"""
├── [21] tk.Canvas (自由グラフィック描画)
│   ├── width / height    : 【影響】描画領域（キャンバス）のピクセルサイズ
│   │                       └─ 例: width=500, height=300
│   ├── bg                : 【影響】キャンバスの背景色を指定
│   │                       └─ 例: bg="#ffffff"
│   ├── scrollregion      : 【影響】可視領域よりも広い描画範囲（スクロール可能範囲）を定義
│   │                       └─ 例: scrollregion=(0, 0, 1000, 1000) # (x1, y1, x2, y2)
│   └── xscrollcommand / yscrollcommand : 【影響】縦横スクロールバーとの連携設定
│                           └─ 例: xscrollcommand=hbar.set
"""
# --- [21] tk.Canvas ---
canvas = tk.Canvas(root, width=200, height=50, bg="white")
canvas.pack(padx=10, pady=5)
canvas.create_rectangle(10, 10, 80, 40, fill="skyblue")
canvas.create_text(130, 25, text="Canvas描画")

"""
└── [22] tk.Message (自動折り返しラベル)
    ├── text              : 【影響】表示する長文メッセージテキスト
    │                       └─ 例: text="ここに長文のメッセージが入ります。"
    ├── width             : 【影響】指定ピクセル幅で強制的に折り返し計算
    │                       └─ 例: width=250
    ├── aspect            : 【影響】幅と高さの縦横比率（%）で自動折り返し形状を決定
    │                       └─ 例: aspect=200 # 幅が高さの2倍（200%）になるよう自動計算
    └── justify           : 【影響】複数行時のテキスト揃え方向
                            └─ 例: justify="center"
"""
# --- [22] tk.Message ---
msg = tk.Message(root, text="これは自動的に折り返される長めのメッセージテキストです。", width=250)
msg.pack(anchor="w", padx=10)


# ==============================================================================
# 7. 補助・レイアウト調整ウィジェット
# ==============================================================================

"""
├── [23] ttk.Scrollbar (スクロールバー)
│   ├── orient            : 【影響】スクロールバーの向き（縦か横か）を指定
│   │                       └─ 例: orient="vertical" # "vertical", "horizontal"
│   └── command           : 【影響】バーを動かしたときにスクロールさせる対象ウィジェットのview関数を指定
│                           └─ 例: command=txt_widget.yview
"""
# --- [23] ttk.Scrollbar ---
# ※ここではTextウィジェットに接続する例
sb = ttk.Scrollbar(root, orient="vertical", command=txt.yview)
txt["yscrollcommand"] = sb.set
sb.pack(side="right", fill="y")

"""
├── [24] ttk.Separator (区切り線)
│   └── orient            : 【影響】画面を区切る枠線の向き（水平線か垂直線か）を指定
│                           └─ 例: orient="horizontal" # "horizontal", "vertical"
"""
# --- [24] ttk.Separator ---
sep = ttk.Separator(root, orient="horizontal")
sep.pack(fill="x", padx=10, pady=10)

"""
└── [25] ttk.Sizegrip (サイズ変更グリップ)
    └── (プロパティ指定はほぼ不要。sizegrip.pack(side="right", anchor="se") で配置するだけで動作)
"""
# --- [25] ttk.Sizegrip ---
sizegrip = ttk.Sizegrip(root)
sizegrip.pack(side="right", anchor="se")


# ==============================================================================
# 8. 対話ダイアログ（単体呼び出し）
# ==============================================================================

def run_dialogs():
    """
    ├── [26] messagebox (ポップアップ)
    │   ├── showinfo(...)     : 【影響】通知用のOKボタンダイアログを表示（アイコン・本文・詳細を指定）
    │   │                       └─ 例: showinfo(title="通知", message="完了しました", detail="処理件数: 10件")
    │   ├── showerror(...)    : 【影響】警告・エラーアイコン付きダイアログを表示
    │   │                       └─ 例: showerror(title="エラー", message="接続失敗")
    │   └── askyesno(...)     : 【影響】「はい/いいえ」の二択ダイアログを表示し返り値（True/False）を得る
    │                           └─ 例: askyesno(title="確認", message="削除しますか？", icon="warning")
    """
    # --- [26] messagebox ---
    messagebox.showinfo("情報", "メッセージボックスの例です")
    ans = messagebox.askyesno("確認", "処理を続行しますか？")
    print(f"Yes/No回答: {ans}")

    """
    ├── [27] filedialog (ファイル選択)
    │   ├── askopenfilename   : 【影響】ファイル選択（開く）画面を表示し、選ばれたパス文字列を得る
    │   │                       └─ 例: askopenfilename(title="開く", initialdir="C:/", filetypes=[("CSV", "*.csv"), ("ALL", "*.*")])
    │   └── asksaveasfilename : 【影響】名前を付けて保存画面を表示し、指定パス文字列を得る
    │                           └─ 例: asksaveasfilename(title="保存", initialfile="result.txt", defaultextension=".txt")
    """
    # --- [27] filedialog ---
    file_path = filedialog.askopenfilename(title="ファイルを選択")
    print(f"選択ファイル: {file_path}")

    """
    └── [28] colorchooser (カラーパレット)
        └── askcolor(...)     : 【影響】OS標準のカラーパレットを表示し、選ばれた色コード（HEX値等）を得る
                                └─ 例: askcolor(title="テーマ色の選択", initialcolor="#ff0000")
    """
    # --- [28] colorchooser ---
    color = colorchooser.askcolor(title="色を選択")
    print(f"選択色(HEX): {color[1]}")





# ダイアログテスト用ボタン
btn_dlg = ttk.Button(root, text="ダイアログ表示テスト", command=run_dialogs)
btn_dlg.pack(pady=5)

root.mainloop()