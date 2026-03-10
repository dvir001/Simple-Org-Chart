"""Tests for simple_org_chart.config – path consistency & helpers."""

from __future__ import annotations

from pathlib import Path

from simple_org_chart.config import (
    BASE_DIR,
    CONFIG_DIR,
    DATA_DIR,
    STATIC_DIR,
    TEMPLATE_DIR,
    SETTINGS_FILE,
    DATA_FILE,
    EMPLOYEE_LIST_FILE,
    as_posix_env,
    ensure_directories,
)


class TestConfigPaths:
    """Verify that config-level path constants are consistent."""

    def test_data_dir_under_base(self):
        assert DATA_DIR.parent == BASE_DIR

    def test_static_dir_under_base(self):
        assert STATIC_DIR.parent == BASE_DIR

    def test_template_dir_under_base(self):
        assert TEMPLATE_DIR.parent == BASE_DIR

    def test_settings_file_in_config_dir(self):
        assert SETTINGS_FILE.parent == CONFIG_DIR

    def test_data_file_in_data_dir(self):
        assert DATA_FILE.parent == DATA_DIR

    def test_employee_list_in_data_dir(self):
        assert EMPLOYEE_LIST_FILE.parent == DATA_DIR


class TestEnsureDirectories:
    def test_creates_directories(self, tmp_path: Path, monkeypatch):
        """ensure_directories should create DATA_DIR, CONFIG_DIR, and STATIC_DIR."""
        import simple_org_chart.config as cfg

        fake_data = tmp_path / "data"
        fake_config = tmp_path / "config"
        fake_static = tmp_path / "static"
        monkeypatch.setattr(cfg, "DATA_DIR", fake_data)
        monkeypatch.setattr(cfg, "CONFIG_DIR", fake_config)
        monkeypatch.setattr(cfg, "STATIC_DIR", fake_static)

        ensure_directories()

        assert fake_data.exists(), "DATA_DIR was not created by ensure_directories()"
        assert fake_config.exists(), "CONFIG_DIR was not created by ensure_directories()"
        assert fake_static.exists(), "STATIC_DIR was not created by ensure_directories()"


class TestAsPosixEnv:
    def test_converts(self):
        mapping = {"DATA": Path("/some/path")}
        result = as_posix_env(mapping)
        assert isinstance(result["DATA"], str)
