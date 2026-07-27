"""
Summary:
    PythonソースコードをAST（抽象構文木）で解析し、仕様情報を抽出するサービスモジュール。
"""

import ast
import os
import re
import typing

import fws_apps.tkinter.fws_spec_generator.models.entity.fws_spec_entities as fws_spec_entities


class FwsSpecParserService:
    """
    Summary:
        Pythonファイルを解析して仕様書用データを生成するサービス。

    Description:
        指定されたディレクトリ配下のPythonファイルを再帰的に探索し、各ファイルの
        クラス、メソッド、単体関数、メンバ変数、および規約に基づくヘッダコメント（docstring）を解析・抽出します。
    """

    def __init__(self) -> None:
        """
        Summary:
            FwsSpecParserServiceクラスのインスタンスを初期化します。
        """
        pass

    def parse_directory(self, dir_path: str) -> typing.List[fws_spec_entities.FwsModuleInfo]:
        """
        Summary:
            指定されたディレクトリ内のすべてのPythonファイルを再帰的に解析します。

        Args:
            self: FwsSpecParserService - FwsSpecParserServiceクラスのインスタンス自身。
            dir_path: str - 解析対象とするディレクトリのパス。

        Returns:
            List[FwsModuleInfo] - 解析されたモジュール（ファイル）情報オブジェクトのリスト。
        """
        modules: typing.List[fws_spec_entities.FwsModuleInfo] = []
        
        if not os.path.exists(dir_path) or not os.path.isdir(dir_path):
            return modules

        for root, _, files in os.walk(dir_path):
            for file in files:
                if file.endswith(".py"):
                    file_path: str = os.path.join(root, file)
                    rel_path: str = os.path.relpath(file_path, dir_path)
                    module_info: typing.Optional[fws_spec_entities.FwsModuleInfo] = self.parse_file(file_path, rel_path)
                    if module_info:
                        modules.append(module_info)
                        
        return modules

    def parse_file(self, file_path: str, relative_path: str) -> typing.Optional[fws_spec_entities.FwsModuleInfo]:
        """
        Summary:
            単一のPythonファイルをASTモジュールで解析します。

        Args:
            self: FwsSpecParserService - FwsSpecParserServiceクラスのインスタンス自身。
            file_path: str - 解析対象とするPythonファイルのパス。
            relative_path: str - 解析対象フォルダからの相対ファイルパス。

        Returns:
            Optional[FwsModuleInfo] - 解析されたモジュール（ファイル）情報。エラー等で解析不能な場合はNone。
        """
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                content: str = f.read()
                
            node: ast.Module = ast.parse(content, filename=file_path)
        except Exception as e:
            print(f"ファイルのパースエラー ({file_path}): {str(e)}")
            return None

        # ファイルヘッダのパース
        module_doc: typing.Optional[str] = ast.get_docstring(node)
        summary, description, args, returns_type, returns_desc, raises, user_actions, screen_name = self._parse_docstring(module_doc)

        # 添付ファイルパスの抽出
        attachment_path: typing.Optional[str] = None
        if module_doc:
            match_att = re.search(r"^\s*Attachment:\s*(.*)$", module_doc, re.MULTILINE)
            if match_att:
                attachment_path = match_att.group(1).strip()

        classes: typing.List[fws_spec_entities.FwsClassInfo] = []
        functions: typing.List[fws_spec_entities.FwsFunctionInfo] = []

        # トップレベルノードの解析
        for child in node.body:
            if isinstance(child, ast.ClassDef):
                class_info: fws_spec_entities.FwsClassInfo = self._parse_class(child, content)
                classes.append(class_info)
            elif isinstance(child, ast.FunctionDef):
                func_info: fws_spec_entities.FwsFunctionInfo = self._parse_function(child, content)
                functions.append(func_info)

        module_name: str = os.path.basename(file_path)

        return fws_spec_entities.FwsModuleInfo(
            file_path=file_path,
            relative_path=relative_path,
            module_name=module_name,
            summary=summary,
            description=description,
            screen_name=screen_name,
            classes=classes,
            functions=functions,
            attachment_path=attachment_path
        )

    def _parse_class(self, class_node: ast.ClassDef, file_content: str) -> fws_spec_entities.FwsClassInfo:
        """
        Summary:
            クラス定義ノードを解析し、クラス情報を抽出します。
        """
        class_doc: typing.Optional[str] = ast.get_docstring(class_node)
        summary, description, _, _, _, _, _, _ = self._parse_docstring(class_doc)

        # メンバ変数（属性）のパース
        attributes: typing.List[fws_spec_entities.FwsVariableInfo] = self._parse_class_attributes(class_node)

        # クラスメソッドのパース
        methods: typing.List[fws_spec_entities.FwsFunctionInfo] = []
        for child in class_node.body:
            if isinstance(child, ast.FunctionDef):
                func_info: fws_spec_entities.FwsFunctionInfo = self._parse_function(child, file_content)
                methods.append(func_info)

        return fws_spec_entities.FwsClassInfo(
            name=class_node.name,
            summary=summary,
            description=description,
            attributes=attributes,
            methods=methods
        )

    def _parse_function(self, func_node: ast.FunctionDef, file_content: str) -> fws_spec_entities.FwsFunctionInfo:
        """
        Summary:
            関数またはメソッド定義ノードを解析します。
        """
        func_doc: typing.Optional[str] = ast.get_docstring(func_node)
        summary, description, args, returns_type, returns_desc, raises, user_actions, _ = self._parse_docstring(func_doc)

        # ドキュメントに記述されている引数と、実際のシグネチャをある程度マージ・補正する
        actual_arg_names: typing.List[str] = [arg.arg for arg in func_node.args.args]
        
        final_args: typing.List[fws_spec_entities.FwsParameterInfo] = []
        for name in actual_arg_names:
            found_arg: typing.Optional[fws_spec_entities.FwsParameterInfo] = None
            for a in args:
                if a.name == name:
                    found_arg = a
                    break
            
            if found_arg:
                final_args.append(found_arg)
            else:
                arg_type: str = "Any"
                for arg_obj in func_node.args.args:
                    if arg_obj.arg == name and arg_obj.annotation:
                        try:
                            arg_type = ast.unparse(arg_obj.annotation)
                        except AttributeError:
                            if isinstance(arg_obj.annotation, ast.Name):
                                arg_type = arg_obj.annotation.id
                            else:
                                arg_type = "Any"
                                
                final_args.append(
                    fws_spec_entities.FwsParameterInfo(
                        name=name,
                        type_name=arg_type,
                        description="（引数の説明がヘッダコメントに記載されていません）"
                    )
                )

        if func_node.returns:
            try:
                actual_ret_type: str = ast.unparse(func_node.returns)
                if actual_ret_type:
                    returns_type = actual_ret_type
            except AttributeError:
                if isinstance(func_node.returns, ast.Name):
                    returns_type = func_node.returns.id

        func_source: str = ""
        try:
            func_source = ast.get_source_segment(file_content, func_node) or ""
        except Exception:
            func_source = "ソースコードの抽出に失敗しました。"

        return fws_spec_entities.FwsFunctionInfo(
            name=func_node.name,
            summary=summary,
            description=description,
            args=final_args,
            returns_type=returns_type if returns_type else "None",
            returns_desc=returns_desc if returns_desc else "戻り値の説明が記載されていません。",
            raises=raises,
            user_actions=user_actions,
            source_code=func_source
        )

    def _parse_class_attributes(self, class_node: ast.ClassDef) -> typing.List[fws_spec_entities.FwsVariableInfo]:
        """
        Summary:
            クラス定義直下の変数、および__init__メソッド内のメンバ変数を抽出し、docstringをパースします。
        """
        attributes: typing.List[fws_spec_entities.FwsVariableInfo] = []
        body_len: int = len(class_node.body)
        
        # 1. クラス定義直下の変数の解析
        for i, child in enumerate(class_node.body):
            is_var: bool = False
            var_name: str = ""
            var_type: str = "Any"
            
            if isinstance(child, ast.AnnAssign):
                if isinstance(child.target, ast.Name):
                    is_var = True
                    var_name = child.target.id
                    try:
                        var_type = ast.unparse(child.annotation)
                    except AttributeError:
                        if isinstance(child.annotation, ast.Name):
                            var_type = child.annotation.id
            elif isinstance(child, ast.Assign):
                if len(child.targets) == 1 and isinstance(child.targets[0], ast.Name):
                    is_var = True
                    var_name = child.targets[0].id

            if is_var:
                var_desc: str = ""
                if i + 1 < body_len:
                    next_node: ast.stmt = class_node.body[i + 1]
                    if isinstance(next_node, ast.Expr) and isinstance(next_node.value, ast.Constant) and isinstance(next_node.value.value, str):
                        raw_desc: str = next_node.value.value.strip()
                        match: typing.Optional[typing.Match[str]] = re.match(r"^([a-zA-Z_0-9\.\答\"'\[\]\s,]+)\s*-\s*(.*)$", raw_desc)
                        if match:
                            var_type = match.group(1).strip()
                            var_desc = match.group(2).strip()
                        else:
                            var_desc = raw_desc

                attributes.append(
                    fws_spec_entities.FwsVariableInfo(
                        name=var_name,
                        type_name=var_type,
                        description=var_desc
                    )
                )

        # 2. コンストラクタ __init__ 内の self.変数名 = 値 の解析
        for child in class_node.body:
            if isinstance(child, ast.FunctionDef) and child.name == "__init__":
                init_len: int = len(child.body)
                for idx, sub in enumerate(child.body):
                    is_self_var: bool = False
                    self_var_name: str = ""
                    self_var_type: str = "Any"
                    
                    if isinstance(sub, ast.Assign):
                        if len(sub.targets) == 1 and isinstance(sub.targets[0], ast.Attribute):
                            attr: ast.Attribute = sub.targets[0]
                            if isinstance(attr.value, ast.Name) and attr.value.id == "self":
                                is_self_var = True
                                self_var_name = attr.attr
                    elif isinstance(sub, ast.AnnAssign):
                        if isinstance(sub.target, ast.Attribute):
                            attr = sub.target
                            if isinstance(attr.value, ast.Name) and attr.value.id == "self":
                                is_self_var = True
                                self_var_name = attr.attr
                                try:
                                    self_var_type = ast.unparse(sub.annotation)
                                except AttributeError:
                                    pass

                    if is_self_var:
                        self_var_desc: str = ""
                        if idx + 1 < init_len:
                            next_sub: ast.stmt = child.body[idx + 1]
                            if isinstance(next_sub, ast.Expr) and isinstance(next_sub.value, ast.Constant) and isinstance(next_sub.value.value, str):
                                raw_desc = next_sub.value.value.strip()
                                match = re.match(r"^([a-zA-Z_0-9\.\答\"'\[\]\s,]+)\s*-\s*(.*)$", raw_desc)
                                if match:
                                    self_var_type = match.group(1).strip()
                                    self_var_desc = match.group(2).strip()
                                else:
                                    self_var_desc = raw_desc
                                    
                        if not any(attr.name == self_var_name for attr in attributes):
                            attributes.append(
                                fws_spec_entities.FwsVariableInfo(
                                    name=self_var_name,
                                    type_name=self_var_type,
                                    description=self_var_desc
                                )
                            )

        return attributes

    def _parse_docstring(
        self,
        doc: typing.Optional[str]
    ) -> typing.Tuple[
        str,
        str,
        typing.List[fws_spec_entities.FwsParameterInfo],
        str,
        str,
        typing.List[fws_spec_entities.FwsExceptionInfo],
        typing.List[fws_spec_entities.FwsUserActionInfo],
        typing.Optional[str]
    ]:
        """
        Summary:
            docstringテキストをセクション（Summary, Description, Args, Returns, Raises, UserAction, ScreenName）に分類・パースします。
        """
        if not doc:
            return "", "", [], "", "", [], [], None

        lines: typing.List[str] = doc.splitlines()
        
        summary_lines: typing.List[str] = []
        desc_lines: typing.List[str] = []
        arg_lines: typing.List[str] = []
        return_lines: typing.List[str] = []
        raise_lines: typing.List[str] = []
        user_action_lines: typing.List[str] = []
        screen_name_lines: typing.List[str] = []
        
        current_state: str = "NONE"
        
        for line in lines:
            stripped: str = line.strip()
            if stripped.startswith("Summary:"):
                current_state = "SUMMARY"
                continue
            elif stripped.startswith("Description:"):
                current_state = "DESC"
                continue
            elif stripped.startswith("Args:"):
                current_state = "ARGS"
                continue
            elif stripped.startswith("Returns:"):
                current_state = "RETURNS"
                continue
            elif stripped.startswith("Raises:"):
                current_state = "RAISES"
                continue
            elif stripped.startswith("UserAction:"):
                current_state = "USER_ACTION"
                continue
            elif stripped.lower().replace(" ", "").startswith("screenname:") or stripped.lower().replace(" ", "").startswith("screenname："):
                current_state = "SCREEN_NAME"
                colon_idx = stripped.find(":")
                if colon_idx == -1:
                    colon_idx = stripped.find("：")
                val = stripped[colon_idx + 1:].strip()
                if val:
                    screen_name_lines.append(val)
                continue
                
            if not stripped:
                continue
                
            if current_state == "SUMMARY":
                summary_lines.append(stripped)
            elif current_state == "DESC":
                desc_lines.append(line)
            elif current_state == "ARGS":
                arg_lines.append(stripped)
            elif current_state == "RETURNS":
                return_lines.append(stripped)
            elif current_state == "RAISES":
                raise_lines.append(stripped)
            elif current_state == "USER_ACTION":
                user_action_lines.append(stripped)
            elif current_state == "SCREEN_NAME":
                screen_name_lines.append(stripped)

        summary: str = " ".join(summary_lines)
        description: str = "\n".join(desc_lines).strip()
        
        args: typing.List[fws_spec_entities.FwsParameterInfo] = []
        for arg_line in arg_lines:
            match: typing.Optional[typing.Match[str]] = re.match(
                r"^([a-zA-Z_]\w*)\s*:\s*([a-zA-Z_0-9\.\答\"'\[\]\s,]+)\s*-\s*(.*)$",
                arg_line
            )
            if match:
                args.append(
                    fws_spec_entities.FwsParameterInfo(
                        name=match.group(1),
                        type_name=match.group(2).strip(),
                        description=match.group(3).strip()
                    )
                )
            else:
                args.append(
                    fws_spec_entities.FwsParameterInfo(
                        name=arg_line,
                        type_name="Any",
                        description=arg_line
                    )
                )
                
        returns_type: str = ""
        returns_desc: str = ""
        if return_lines:
            ret_line: str = return_lines[0]
            match_ret: typing.Optional[typing.Match[str]] = re.match(
                r"^([a-zA-Z_0-9\.\答\"'\[\]\s,]+)\s*-\s*(.*)$",
                ret_line
            )
            if match_ret:
                returns_type = match_ret.group(1).strip()
                returns_desc = match_ret.group(2).strip()
            else:
                returns_desc = ret_line

        raises: typing.List[fws_spec_entities.FwsExceptionInfo] = []
        for raise_line in raise_lines:
            match_raise: typing.Optional[typing.Match[str]] = re.match(r"^([a-zA-Z_0-9\.]+)\s*-\s*(.*)$", raise_line)
            if match_raise:
                raises.append(
                    fws_spec_entities.FwsExceptionInfo(
                        name=match_raise.group(1).strip(),
                        description=match_raise.group(2).strip()
                    )
                )
            else:
                raises.append(
                    fws_spec_entities.FwsExceptionInfo(
                        name=raise_line,
                        description=raise_line
                    )
                )

        user_actions: typing.List[fws_spec_entities.FwsUserActionInfo] = []
        for ua_line in user_action_lines:
            match_ua: typing.Optional[typing.Match[str]] = re.match(r"^(.*?)\s*-\s*(.*)$", ua_line)
            if match_ua:
                user_actions.append(
                    fws_spec_entities.FwsUserActionInfo(
                        trigger=match_ua.group(1).strip(),
                        action=match_ua.group(2).strip()
                    )
                )
            else:
                user_actions.append(
                    fws_spec_entities.FwsUserActionInfo(
                        trigger=ua_line,
                        action=""
                    )
                )

        screen_name: typing.Optional[str] = " ".join(screen_name_lines).strip() or None

        return summary, description, args, returns_type, returns_desc, raises, user_actions, screen_name
