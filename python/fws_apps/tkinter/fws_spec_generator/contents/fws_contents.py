"""
Summary:
    仕様書自動生成ツールで用いる静的HTML/CSSテンプレートおよびデザイン定数を定義するモジュール。
"""

# HTML 骨組みテンプレート
HTML_TEMPLATE: str = """<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>API Specifications (Javadoc Style)</title>
    <style>
        {style_content}
    </style>
</head>
<body>
    <div class="sidebar">
        <div class="sidebar-header">API Specifications</div>
        <div class="sidebar-menu">
            {sidebar_menu}
        </div>
    </div>
    <div class="resizer" id="dragMe"></div>
    <div class="content">
        {user_guide_body}
        {content_body}
    </div>
    <script>
        document.addEventListener("DOMContentLoaded", function() {{
            const modules = document.querySelectorAll(".module-section");
            
            function showModuleFromHash() {{
                const hash = window.location.hash;
                if (!hash) {{
                    // ハッシュが指定されていない場合はユーザーガイドを表示
                    window.location.hash = "#user_guide";
                    return;
                }}

                // ユーザーガイド全体、または個別操作仕様トリガーへのジャンプの場合
                if (hash === "#user_guide" || hash.startsWith("#user_guide_") || hash.startsWith("#ua_")) {{
                    let activeGuideId = "user_guide";
                    if (hash.startsWith("#user_guide_")) {{
                        activeGuideId = hash.substring(1);
                    }} else if (hash.startsWith("#ua_")) {{
                        // #ua_method_0_1_2_0 -> parts = ["ua", "method", "0", "1", "2", "0"]
                        // parts[1] (モジュールインデックス "0") を取得して対応する画面用ユーザーガイドに切り替える
                        const parts = hash.substring(4).split("_");
                        if (parts.length >= 2) {{
                            const modIndex = parts[1];
                            activeGuideId = "user_guide_" + modIndex;
                        }}
                    }}

                    // 該当するユーザーガイドページのみをアクティブ表示にする
                    const userGuides = document.querySelectorAll(".user-guide-section");
                    userGuides.forEach(g => {{
                        if (g.id === activeGuideId) {{
                            g.style.display = "block";
                        }} else {{
                            g.style.display = "none";
                        }}
                    }});

                    // 開発者用ドキュメントはすべて非表示にする
                    modules.forEach(m => {{
                        m.style.display = "none";
                    }});

                    // ユーザーガイド一覧表内の該当行へスムーズスクロール
                    const targetElement = document.getElementById(hash.substring(1));
                    if (targetElement) {{
                        setTimeout(() => {{
                            targetElement.scrollIntoView({{ behavior: "smooth", block: "start" }});
                        }}, 50);
                    }}
                    return;
                }}
                
                // 開発者ドキュメント閲覧時はすべてのユーザーガイド表示領域を非表示にする
                const userGuides = document.querySelectorAll(".user-guide-section");
                userGuides.forEach(g => {{
                    g.style.display = "none";
                }});

                // ハッシュ文字列（例: #class_0_1, #mod_0, #func_0_2）からモジュールインデックスを解析
                const parts = hash.substring(1).split("_");
                if (parts.length >= 2) {{
                    const modIndex = parts[1];
                    const targetModId = "mod_" + modIndex;
                    
                    modules.forEach(m => {{
                        if (m.id === targetModId) {{
                            m.style.display = "block";
                        }} else {{
                            m.style.display = "none";
                        }}
                    }});

                    // 該当要素へスムーズにスクロール
                    const targetElement = document.getElementById(hash.substring(1));
                    if (targetElement) {{
                        setTimeout(() => {{
                            targetElement.scrollIntoView({{ behavior: "smooth", block: "start" }});
                        }}, 50);
                    }}
                }}
            }}

            window.addEventListener("hashchange", showModuleFromHash);
            showModuleFromHash(); // 初期ロード時実行

            // サイドバーリサイズ機能の実装
            const resizer = document.getElementById("dragMe");
            const sidebar = document.querySelector(".sidebar");
            let isResizing = false;

            resizer.addEventListener("mousedown", function(e) {{
                isResizing = true;
                document.body.style.cursor = "col-resize";
                document.body.style.userSelect = "none";
            }});

            document.addEventListener("mousemove", function(e) {{
                if (!isResizing) return;
                
                let width = e.clientX;
                if (width < 150) {{
                    width = 150;
                }}
                if (width > 600) {{
                    width = 600;
                }}
                
                sidebar.style.width = width + "px";
            }});

            document.addEventListener("mouseup", function() {{
                if (isResizing) {{
                    isResizing = false;
                    document.body.style.cursor = "default";
                    document.body.style.userSelect = "auto";
                }}
            }});
        }});
    </script>
</body>
</html>
"""

# CSS スタイルシートテンプレート (Pythonでのformat展開を行わないため、中括弧は一重で定義します)
CSS_TEMPLATE: str = """
        body {
            margin: 0;
            padding: 0;
            font-family: "Segoe UI", Meiryo, sans-serif;
            background-color: #1e1e1e;
            color: #d4d4d4;
            display: flex;
            height: 100vh;
            overflow: hidden;
        }
        .font-code {
            font-family: Consolas, "Courier New", monospace;
        }
        .sidebar {
            width: 300px;
            background-color: #252526;
            border-right: 1px solid #3c3c3c;
            display: flex;
            flex-direction: column;
            flex-shrink: 0;
        }
        .sidebar-header {
            padding: 20px;
            font-size: 18px;
            font-weight: bold;
            color: #569cd6;
            border-bottom: 1px solid #3c3c3c;
        }
        .sidebar-menu {
            flex: 1;
            overflow-y: auto;
            padding: 15px 10px;
        }
        details.sidebar-folder {
            margin: 4px 0 4px 10px;
        }
        details.sidebar-folder[open] > summary {
            color: #569cd6;
        }
        summary.sidebar-summary {
            cursor: pointer;
            font-size: 13px;
            color: #cccccc;
            padding: 4px 8px;
            outline: none;
            user-select: none;
            transition: color 0.2s ease;
        }
        summary.sidebar-summary:hover {
            color: #ffffff;
        }
        .sidebar-folder-content {
            border-left: 1px dashed #3c3c3c;
            margin-left: 12px;
            padding-left: 4px;
        }
        .menu-item {
            padding: 6px 15px;
            cursor: pointer;
            display: block;
            color: #cccccc;
            text-decoration: none;
            font-size: 13px;
            border-left: 2px solid transparent;
            transition: all 0.2s ease;
        }
        .menu-item:hover {
            background-color: #2d2d2d;
            color: #ffffff;
            border-left-color: #007acc;
        }
        .menu-sub-item {
            padding: 4px 15px 4px 28px;
            font-size: 12px;
            color: #9cdcfe;
        }
        .menu-sub-item.font-func {
            color: #dcdcaa;
        }
        .content {
            flex: 1;
            overflow-y: auto;
            padding: 40px 60px;
            scroll-behavior: smooth;
        }
        .module-section {
            margin-bottom: 80px;
            border-bottom: 2px solid #3c3c3c;
            padding-bottom: 40px;
            display: none; /* 初期状態は非表示（JavaScriptで動的に切り替えます） */
        }
        .module-title {
            font-size: 26px;
            color: #569cd6;
            border-bottom: 1px solid #3c3c3c;
            padding-bottom: 8px;
            margin-top: 0;
            margin-bottom: 15px;
        }
        .class-section {
            background-color: #252526;
            border: 1px solid #3c3c3c;
            border-radius: 6px;
            padding: 25px;
            margin-bottom: 40px;
        }
        .class-title {
            font-size: 20px;
            color: #4ec9b0;
            margin-top: 0;
            margin-bottom: 15px;
            border-bottom: 1px solid #3c3c3c;
            padding-bottom: 6px;
        }
        .func-section {
            background-color: #2d2d2d;
            border-left: 4px solid #007acc;
            border-radius: 0 4px 4px 0;
            padding: 20px;
            margin-bottom: 30px;
        }
        .func-title {
            font-size: 16px;
            color: #dcdcaa;
            margin-top: 0;
            margin-bottom: 10px;
        }
        .section-desc {
            font-size: 14px;
            line-height: 1.6;
            color: #cccccc;
            margin-bottom: 20px;
            white-space: pre-wrap;
        }
        .table-title {
            font-size: 13px;
            font-weight: bold;
            margin-top: 15px;
            margin-bottom: 6px;
            color: #569cd6;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 20px;
            font-size: 13px;
            background-color: #1e1e1e;
        }
        th, td {
            border: 1px solid #3c3c3c;
            padding: 8px 12px;
            text-align: left;
            vertical-align: top;
        }
        th {
            background-color: #2d2d2d;
            color: #569cd6;
            width: 25%;
        }
        td.type-col {
            color: #4ec9b0;
            width: 20%;
        }
        td.name-col {
            color: #9cdcfe;
            width: 20%;
        }
        .badge {
            display: inline-block;
            padding: 2px 6px;
            font-size: 11px;
            font-weight: bold;
            border-radius: 3px;
            margin-right: 8px;
            text-transform: uppercase;
        }
        .badge-module { background-color: #007acc; color: white; }
        .badge-class { background-color: #4ec9b0; color: black; }
        .badge-func { background-color: #dcdcaa; color: black; }

        /* クラス内のアコーディオン */
        details.sidebar-class-folder {
            margin: 2px 0 2px 15px;
        }
        details.sidebar-class-folder[open] > summary {
            color: #4ec9b0;
        }
        summary.sidebar-class-summary {
            cursor: pointer;
            font-size: 12px;
            color: #a3a3a3;
            padding: 3px 6px;
            outline: none;
            user-select: none;
            transition: color 0.2s ease;
        }
        summary.sidebar-class-summary:hover {
            color: #ffffff;
        }
        .sidebar-class-content {
            border-left: 1px dashed #444444;
            margin-left: 10px;
            padding-left: 2px;
        }
        .menu-item-link {
            color: inherit;
            text-decoration: none;
        }
        .menu-item-link:hover {
            color: #ffffff;
        }

        /* 実ソースコードの表示用スタイル */
        .source-code-details {
            margin-top: 15px;
            border: 1px solid #3c3c3c;
            border-radius: 4px;
            background-color: #2d2d2d;
            overflow: hidden;
        }
        .source-code-summary {
            cursor: pointer;
            padding: 8px 12px;
            font-size: 12px;
            font-weight: bold;
            background-color: #252526;
            color: #cccccc;
            user-select: none;
            outline: none;
            transition: background-color 0.2s ease, color 0.2s ease;
        }
        .source-code-summary:hover {
            background-color: #2d2d2d;
            color: #ffffff;
        }
        .source-code-details[open] .source-code-summary {
            border-bottom: 1px solid #3c3c3c;
        }
        
        /* 自前での行番号付きコードスニペットのスタイル定義 */
        .code-snippet {
            display: flex;
            background-color: #1e1e1e;
            font-family: Consolas, "Courier New", monospace;
            font-size: 12.5px;
            line-height: 1.5;
            max-height: 400px;
            overflow-y: auto;
        }
        .code-line-numbers {
            padding: 15px 10px;
            background-color: #1e1e1e;
            color: #858585;
            border-right: 1px solid #3c3c3c;
            text-align: right;
            user-select: none;
            white-space: pre;
            min-width: 25px;
        }
        .code-body-pre {
            margin: 0;
            padding: 15px;
            flex: 1;
            align-self: flex-start;
            overflow-x: auto;
            overflow-y: hidden;
            background-color: transparent;
        }
        .code-body-pre code {
            font-family: inherit;
            color: #d4d4d4;
            white-space: pre;
        }
        
        /* 添付ファイルエリア of スタイル */
        .attachment-box {
            margin-top: 15px;
            margin-bottom: 15px;
            padding: 12px 18px;
            background-color: #252526;
            border: 1px solid #3c3c3c;
            border-radius: 4px;
            display: inline-flex;
            align-items: center;
            font-size: 13px;
        }
        .attachment-label {
            font-weight: bold;
            color: #4ec9b0;
            margin-right: 10px;
        }
        .attachment-link {
            color: #569cd6;
            text-decoration: none;
            border-bottom: 1px dashed #569cd6;
            transition: color 0.2s ease, border-color 0.2s ease;
        }
        .attachment-link:hover {
            color: #9cdcfe;
            border-bottom-style: solid;
        }

        /* スプリッター（レザイザー） of スタイル */
        .resizer {
            width: 4px;
            cursor: col-resize;
            background-color: #3c3c3c;
            transition: background-color 0.2s ease;
            flex-shrink: 0;
        }
        .resizer:hover, .resizer:active {
            background-color: #007acc;
        }

        /* ユーザーガイド関連 of スタイル */
        .user-guide-section {
            margin-bottom: 80px;
            border-bottom: 2px solid #3c3c3c;
            padding-bottom: 40px;
            display: none;
        }
        .user-guide-title {
            font-size: 26px;
            color: #4ec9b0;
            border-bottom: 1px solid #3c3c3c;
            padding-bottom: 8px;
            margin-top: 0;
            margin-bottom: 15px;
        }
        .guide-table th {
            background-color: #2d2d2d;
            color: #4ec9b0;
        }
        .guide-table th:nth-child(1) { width: 30%; }
        .guide-table th:nth-child(2) { width: 50%; }
        .guide-table th:nth-child(3) { width: 20%; }
"""
