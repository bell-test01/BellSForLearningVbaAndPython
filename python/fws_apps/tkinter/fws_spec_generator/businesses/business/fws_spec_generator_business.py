"""解析中"""
"""
Summary:
    解析されたソースコード仕様データからHTML仕様書を生成するサービスモジュール。
"""

import html as html_module
import os
import typing

from fws_apps.tkinter.fws_spec_generator.businesses.entity import fws_spec_module_info
from fws_apps.tkinter.fws_spec_generator.businesses.entity import fws_spec_class_info
from fws_apps.tkinter.fws_spec_generator.businesses.entity import fws_spec_function_info
from fws_apps.tkinter.fws_spec_generator.businesses.entity import fws_spec_parameter_info
from fws_apps.tkinter.fws_spec_generator.businesses.entity import fws_spec_variable_info
from fws_apps.tkinter.fws_spec_generator.contents import fws_contents as fws_contents


class FwsSpecGeneratorBusiness:
    """
    Summary:
        仕様情報をもとにHTML仕様書を生成するジェネレータサービス。
    """

    def generate_html_spec(self,modules: typing.List[fws_spec_module_info.FwsSpecModuleInfo], output_path: str) -> str:
        """
        Summary:
            解析されたモジュール情報のリストからHTML仕様書を生成し、指定パスへ出力します。

        Args:
            self: FwsSpecGeneratorBusiness - FwsSpecGeneratorBusinessクラスのインスタンス自身。
            modules: List[FwsModuleInfo] - 解析済みの全モジュール仕様情報リスト。
            output_path: str - 生成されるHTML仕様書のファイル出力先パス。

        Returns:
            str - 生成された仕様書HTMLファイルの絶対パス。
        """
        html_content = self._build_html(modules)
        abs_path = os.path.abspath(output_path)
        os.makedirs(os.path.dirname(abs_path), exist_ok=True)
        with open(abs_path, "w", encoding="utf-8") as f:
            f.write(html_content)
        return abs_path

    def _build_html(self, modules: typing.List[fws_spec_module_info.FwsSpecModuleInfo]) -> str:
        """
        Summary:
            モジュール情報のリストからHTML全体の文字列を構築します。
        """
        # --- 1. ユーザーガイド用サイドバーツリーの構築 ---
        user_guide_menu = '<details class="sidebar-folder">'
        user_guide_menu += '<summary class="sidebar-summary" style="font-weight: bold; color: #4ec9b0;"><a href="#user_guide" class="menu-item-link">📁 ユーザーガイド</a></summary>'
        user_guide_menu += '<div class="sidebar-folder-content">'
        
        has_any_user_action = False
        for m_idx, m in enumerate(modules):
            module_actions = []
            
            # 単体関数のチェック
            for f_idx, f in enumerate(m.functions):
                if f.user_actions:
                    f_id = f"func_{m_idx}_{f_idx}"
                    module_actions.append((f_id, f.user_actions))
                    
            # クラスメソッドのチェック
            for c_idx, c in enumerate(m.classes):
                for meth_idx, meth in enumerate(c.methods):
                    if meth.user_actions:
                        meth_id = f"method_{m_idx}_{c_idx}_{meth_idx}"
                        module_actions.append((meth_id, meth.user_actions))
                        
            if module_actions:
                has_any_user_action = True
                display_name = m.screen_name if m.screen_name else m.module_name
                user_guide_menu += '<details class="sidebar-folder">'
                user_guide_menu += f'<summary class="sidebar-summary font-code"><a href="#user_guide_{m_idx}" class="menu-item-link">📁 {display_name}</a></summary>'
                user_guide_menu += '<div class="sidebar-folder-content">'
                for target_id, actions in module_actions:
                    for act_idx, act in enumerate(actions):
                        ua_id = f"ua_{target_id}_{act_idx}"
                        user_guide_menu += f'<a href="#{ua_id}" class="menu-item font-code font-func">🖱️ {act.trigger}</a>'
                user_guide_menu += '</div></details>'
                
        if not has_any_user_action:
            user_guide_menu += '<div class="menu-item" style="font-style: italic; color: #888888;">操作仕様が定義されていません</div>'
            
        user_guide_menu += '</div></details>'

        # --- 2. デベロッパガイド（API仕様）用サイドバーツリーの構築 ---
        tree: typing.Dict[str, typing.Any] = {}
        for m_idx, m in enumerate(modules):
            parts = m.relative_path.replace("\\", "/").split("/")
            current = tree
            for part in parts[:-1]:
                if part not in current:
                    current[part] = {}
                current = current[part]
            if "__files__" not in current:
                current["__files__"] = []
            current["__files__"].append((parts[-1], m, m_idx))

        def render_tree(node: typing.Dict[str, typing.Any]) -> str:
            html = ""
            for key, val in node.items():
                if key == "__files__":
                    for fname, m_info, m_index in val:
                        m_id = f"mod_{m_index}"
                        html += f'<details class="sidebar-folder">'
                        html += f'<summary class="sidebar-summary font-code"><a href="#{m_id}" class="menu-item-link">📄 {m_info.module_name}</a></summary>'
                        html += f'<div class="sidebar-folder-content">'
                        
                        for c_idx, c in enumerate(m_info.classes):
                            c_id = f"class_{m_index}_{c_idx}"
                            html += f'<details class="sidebar-class-folder">'
                            html += f'<summary class="sidebar-class-summary font-code"><a href="#{c_id}" class="menu-item-link">class {c.name}</a></summary>'
                            html += f'<div class="sidebar-class-content">'
                            
                            for a_idx, a in enumerate(c.attributes):
                                a_id = f"attr_{m_index}_{c_idx}_{a_idx}"
                                html += f'<a href="#{a_id}" class="menu-item menu-sub-item font-code">var {a.name}</a>'
                                
                            for meth_idx, meth in enumerate(c.methods):
                                meth_id = f"method_{m_index}_{c_idx}_{meth_idx}"
                                html += f'<a href="#{meth_id}" class="menu-item menu-sub-item font-code font-func">def {meth.name}</a>'
                                
                            html += '</div></details>'
                        
                        for f_idx, f in enumerate(m_info.functions):
                            f_id = f"func_{m_index}_{f_idx}"
                            html += f'<a href="#{f_id}" class="menu-item menu-sub-item font-code font-func">def {f.name}</a>'
                            
                        html += '</div></details>'
                else:
                    html += f'<details class="sidebar-folder">'
                    html += f'<summary class="sidebar-summary">📁 {key}</summary>'
                    html += f'<div class="sidebar-folder-content">'
                    html += render_tree(val)
                    html += '</div></details>'
            return html

        developer_menu = '<details class="sidebar-folder">'
        developer_menu += '<summary class="sidebar-summary" style="font-weight: bold; color: #569cd6;">📁 デベロッパガイド</summary>'
        developer_menu += '<div class="sidebar-folder-content">'
        developer_menu += render_tree(tree)
        developer_menu += '</div></details>'

        # 2つのメニューをサイドバーに統合
        sidebar_menu = user_guide_menu + developer_menu
        
        # 画面別のユーザーガイドHTML構築および結合
        user_guide_body = '<div id="user_guide" class="user-guide-section">'
        user_guide_body += '<h1 class="user-guide-title">📁 ユーザーガイド</h1>'
        user_guide_body += '<div class="section-desc">左サイドバーのメニューから、各画面の操作仕様やトリガーを選択してください。</div>'
        user_guide_body += '</div>'
        for m_idx, m in enumerate(modules):
            user_guide_body += self._build_module_user_guide(m, m_idx)
            
        content_body: str = ""

        # 右コンテンツエリアの組み立て
        for m_idx, m in enumerate(modules):
            m_id: str = f"mod_{m_idx}"
            content_body += self._build_module_content(m, m_id, m_idx)

        # HTML全体の組み立て
        html_template: str = fws_contents.HTML_TEMPLATE
        style_content: str = fws_contents.CSS_TEMPLATE

        return html_template.format(
            style_content=style_content,
            sidebar_menu=sidebar_menu,
            user_guide_body=user_guide_body,
            content_body=content_body
        )

    def _build_module_user_guide(self, m: fws_spec_module_info.FwsSpecModuleInfo, m_idx: int) -> str:
        """
        Summary:
            特定のモジュール（画面）に特化したユーザーガイドHTMLを構築します。
        """
        display_name = m.screen_name if m.screen_name else m.module_name
        html = f'<div id="user_guide_{m_idx}" class="user-guide-section">'
        html += f'<h1 class="user-guide-title">📁 {display_name} (操作仕様)</h1>'
        html += f'<div class="section-desc">{display_name}における画面上の操作トリガーと、それに対応して期待される内部動作仕様の一覧を確認することができます。</div>'
        
        has_any_action = False
        table_html = '<table class="guide-table"><thead><tr><th>操作対象（トリガー）</th><th>期待される動作仕様</th><th>定義箇所 (モジュール / メンバ)</th></tr></thead><tbody>'
        
        # 単体関数
        for f_idx, f in enumerate(m.functions):
            if f.user_actions:
                has_any_action = True
                f_id = f"func_{m_idx}_{f_idx}"
                loc_html = f'<a href="#{f_id}" class="attachment-link" style="font-size: 11px; font-weight: bold;">{m.module_name}<br><span style="font-size: 10px; font-weight: normal; color: #dcdcaa;">def {f.name}</span></a>'
                for act_idx, act in enumerate(f.user_actions):
                    ua_id = f"ua_{f_id}_{act_idx}"
                    table_html += f'<tr id="{ua_id}"><td class="name-col font-code">{html_module.escape(act.trigger)}</td><td>{html_module.escape(act.action)}</td><td class="font-code">{loc_html}</td></tr>'
                    
        # クラスメソッド
        for c_idx, c in enumerate(m.classes):
            for meth_idx, meth in enumerate(c.methods):
                if meth.user_actions:
                    has_any_action = True
                    meth_id = f"method_{m_idx}_{c_idx}_{meth_idx}"
                    loc_html = f'<a href="#{meth_id}" class="attachment-link" style="font-size: 11px; font-weight: bold;">{m.module_name}<br><span style="font-size: 10px; font-weight: normal; color: #4ec9b0;">class {c.name}</span><br><span style="font-size: 10px; font-weight: normal; color: #dcdcaa;">def {meth.name}</span></a>'
                    for act_idx, act in enumerate(meth.user_actions):
                        ua_id = f"ua_{meth_id}_{act_idx}"
                        table_html += f'<tr id="{ua_id}"><td class="name-col font-code">{html_module.escape(act.trigger)}</td><td>{html_module.escape(act.action)}</td><td class="font-code">{loc_html}</td></tr>'
                        
        table_html += '</tbody></table>'
        
        if has_any_action:
            html += table_html
        else:
            html += '<div class="section-desc" style="font-style: italic;">この画面には操作仕様が定義されていません。</div>'
            
        html += '</div>'
        return html

    def _build_module_content(self, m: fws_spec_module_info.FwsSpecModuleInfo, m_id: str, m_idx: int) -> str:
        """
        Summary:
            モジュール（ファイル）ごとの仕様ドキュメント部分のHTMLを構築します。
        """
        html = f'<div id="{m_id}" class="module-section">'
        html += f'<h1 class="module-title"><span class="badge badge-module">module</span>{m.module_name}</h1>'
        html += f'<div class="section-desc">{html_module.escape(m.summary)}\n\n{html_module.escape(m.description)}</div>'

        if m.attachment_path:
            basename = os.path.basename(m.attachment_path)
            file_uri = "file:///" + os.path.abspath(m.attachment_path).replace("\\", "/")
            html += f"""
            <div class="attachment-box">
                <span class="attachment-label">添付設計書:</span>
                <a href="{file_uri}" class="attachment-link" target="_blank">{basename}</a>
            </div>
            """

        for c_idx, c in enumerate(m.classes):
            c_id = f"class_{m_idx}_{c_idx}"
            html += f'<div id="{c_id}" class="class-section">'
            html += f'<h2 class="class-title"><span class="badge badge-class">class</span>{c.name}</h2>'
            html += f'<div class="section-desc">{html_module.escape(c.summary)}\n\n{html_module.escape(c.description)}</div>'

            if c.attributes:
                html += '<div class="table-title">属性（フィールド）</div>'
                html += '<table><thead><tr><th>名前</th><th>型</th><th>説明</th></tr></thead><tbody>'
                for a in c.attributes:
                    a_id = f"attr_{m_idx}_{c_idx}_{c.attributes.index(a)}"
                    html += f'<tr id="{a_id}"><td class="name-col font-code">{a.name}</td><td class="type-col font-code">{a.type_name}</td><td>{html_module.escape(a.description)}</td></tr>'
                html += '</tbody></table>'

            for meth_idx, meth in enumerate(c.methods):
                meth_id = f"method_{m_idx}_{c_idx}_{meth_idx}"
                html += self._build_function_content(meth, meth_id, is_method=True)

            html += '</div>'

        for f_idx, f in enumerate(m.functions):
            f_id = f"func_{m_idx}_{f_idx}"
            html += self._build_function_content(f, f_id, is_method=False)

        html += '</div>'
        return html

    def _build_function_content(self, f: fws_spec_function_info.FwsSpecFunctionInfo, f_id: str, is_method: bool) -> str:
        """
        Summary:
            関数またはメソッドごとの仕様ドキュメント部分のHTMLを構築します。
        """
        badge = "method" if is_method else "def"
        badge_cls = "badge-func"
        
        html = f'<div id="{f_id}" class="func-section">'
        html += f'<h3 class="func-title"><span class="badge {badge_cls}">{badge}</span>{f.name}</h3>'
        html += f'<div class="section-desc">{html_module.escape(f.summary)}\n\n{html_module.escape(f.description)}</div>'

        if f.args:
            html += '<div class="table-title">引数</div>'
            html += '<table><thead><tr><th>引数名</th><th>データ型</th><th>説明</th></tr></thead><tbody>'
            for arg in f.args:
                html += f'<tr><td class="name-col font-code">{arg.name}</td><td class="type-col font-code">{arg.type_name}</td><td>{html_module.escape(arg.description)}</td></tr>'
            html += '</tbody></table>'

        if f.returns:
            html += '<div class="table-title">戻り値</div>'
            html += '<table><thead><tr><th>データ型</th><th>説明</th></tr></thead><tbody>'
            for return_info in f.returns:
                html += f'<tr><td class="type-col font-code">{return_info.type_name}</td><td>{html_module.escape(return_info.description)}</td></tr>'
            html += '</tbody></table>'

        if f.raises:
            html += '<div class="table-title">発生しうる例外</div>'
            html += '<table><thead><tr><th>例外クラス</th><th>発生条件</th></tr></thead><tbody>'
            for raise_info in f.raises:
                html += f'<tr><td class="name-col font-code">{raise_info.name}</td><td>{html_module.escape(raise_info.description)}</td></tr>'
            html += '</tbody></table>'

        if f.user_actions:
            html += '<div class="table-title">操作仕様 (UserAction)</div>'
            html += '<table><thead><tr><th>操作対象（トリガー）</th><th>動作仕様の説明</th></tr></thead><tbody>'
            for act_idx, act in enumerate(f.user_actions):
                detail_ua_id = f"detail_ua_{f_id}_{act_idx}"
                html += f'<tr id="{detail_ua_id}"><td class="name-col font-code">{html_module.escape(act.trigger)}</td><td>{html_module.escape(act.action)}</td></tr>'
            html += '</tbody></table>'

        if f.source_code:
            lines = f.source_code.splitlines()
            numbers_txt = "\n".join(str(i) for i in range(1, len(lines) + 1))
            html += f"""
            <details class="source-code-details">
                <summary class="source-code-summary">ソースコード表示</summary>
                <div class="code-snippet">
                    <div class="code-line-numbers">{numbers_txt}</div>
                    <pre class="code-body-pre"><code>{html_module.escape(f.source_code)}</code></pre>
                </div>
            </details>
            """

        html += '</div>'
        return html
