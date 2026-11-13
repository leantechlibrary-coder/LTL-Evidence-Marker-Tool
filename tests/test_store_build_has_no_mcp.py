"""Microsoft Store 版（GUI 専用）に MCP が紛れ込まないことの歯止め。

Store 版の exe は entry.py を起点に PyInstaller が import を辿って作る。
entry.py と、そこから読み込まれる自作モジュールが mcp / mcp_server を
import していなければ、MCP は exe に入らない。
"""
import ast
import os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
STORE_MODULES = ["entry.py", "pdf_evidence_marker.py", "stamp_core.py"]
FORBIDDEN = {"mcp", "mcp_server"}


def _imported_top_level_names(path):
    with open(path, encoding="utf-8") as f:
        tree = ast.parse(f.read())
    names = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            names.update(a.name.split(".")[0] for a in node.names)
        elif isinstance(node, ast.ImportFrom) and node.module:
            names.add(node.module.split(".")[0])
    return names


def test_store_modules_do_not_import_mcp():
    for name in STORE_MODULES:
        bad = _imported_top_level_names(os.path.join(ROOT, name)) & FORBIDDEN
        assert not bad, f"{name} が {sorted(bad)} を import しています（Store 版に MCP が入ります）"


def test_store_manifest_has_no_execution_alias():
    with open(os.path.join(ROOT, "packaging", "AppxManifest.xml"), encoding="utf-8") as f:
        manifest = f.read()
    assert "appExecutionAlias" not in manifest
