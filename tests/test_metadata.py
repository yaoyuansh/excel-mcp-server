"""Package metadata consistency tests."""

import tomllib
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PYPROJECT = ROOT / "pyproject.toml"


def _project():
    return tomllib.loads(PYPROJECT.read_text(encoding="utf-8"))["project"]


def test_license_metadata_is_apache_2():
    project = _project()
    assert project["license"] == "Apache-2.0"
    assert "License :: OSI Approved :: Apache Software License" in project["classifiers"]
    assert "License :: OSI Approved :: MIT License" not in project["classifiers"]
    assert (ROOT / "LICENSE").read_text(encoding="utf-8").lstrip().startswith("Apache License")


def test_console_script_entrypoint_exists():
    project_data = tomllib.loads(PYPROJECT.read_text(encoding="utf-8"))
    assert project_data["project"]["scripts"]["excel-mcp-server"] == "excel_mcp.server:main"


def test_core_dependencies_declared():
    dependencies = _project()["dependencies"]
    assert "mcp>=1.0.0" in dependencies
    assert "pywin32>=306" in dependencies
    assert "Pillow>=10.0.0" in dependencies
