"""Regression tests for accidental JSON literals in Python code."""

import ast
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
BAD_NAMES = {"true", "false", "null"}


def test_no_lowercase_json_literals_as_python_names():
    offenders = []
    for path in SRC.rglob("*.py"):
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        for node in ast.walk(tree):
            if isinstance(node, ast.Name) and node.id in BAD_NAMES:
                offenders.append(f"{path.relative_to(ROOT)}:{node.lineno}:{node.col_offset}:{node.id}")

    assert offenders == []
