"""Tests for utility functions in mcp_server.py."""

import os
import json

import pytest
from pptx.dml.color import RGBColor

from onit_office.mcp_server import (
    _parse_color,
    _resolve_data_path,
    _secure_makedirs,
    _ensure_directory,
    _read_file_as_base64,
)


class TestParseColor:
    """Tests for _parse_color()."""

    def test_named_colors(self):
        assert _parse_color("red") == RGBColor(255, 0, 0)
        assert _parse_color("green") == RGBColor(0, 255, 0)
        assert _parse_color("blue") == RGBColor(0, 0, 255)
        assert _parse_color("white") == RGBColor(255, 255, 255)
        assert _parse_color("black") == RGBColor(0, 0, 0)
        assert _parse_color("yellow") == RGBColor(255, 255, 0)
        assert _parse_color("orange") == RGBColor(255, 165, 0)
        assert _parse_color("purple") == RGBColor(128, 0, 128)

    def test_gray_grey_alias(self):
        assert _parse_color("gray") == RGBColor(128, 128, 128)
        assert _parse_color("grey") == RGBColor(128, 128, 128)

    def test_case_insensitive(self):
        assert _parse_color("RED") == RGBColor(255, 0, 0)
        assert _parse_color("Blue") == RGBColor(0, 0, 255)

    def test_hex_color(self):
        assert _parse_color("#FF0000") == RGBColor(255, 0, 0)
        assert _parse_color("#00FF00") == RGBColor(0, 255, 0)
        assert _parse_color("#4472C4") == RGBColor(68, 114, 196)

    def test_invalid_color_returns_black(self):
        assert _parse_color("nonexistent") == RGBColor(0, 0, 0)
        assert _parse_color("#ZZZZZZ") == RGBColor(0, 0, 0)
        assert _parse_color("#FF") == RGBColor(0, 0, 0)  # too short


class TestResolveDataPath:
    """Tests for _resolve_data_path()."""

    def test_relative_path_placed_under_data_path(self, tmp_data_dir):
        result = _resolve_data_path("myfile.pptx")
        assert result == os.path.join(os.path.abspath(str(tmp_data_dir)), "myfile.pptx")

    def test_absolute_path_outside_data_path(self, tmp_data_dir):
        result = _resolve_data_path("/some/other/path/file.pptx")
        assert result.endswith("file.pptx")
        assert os.path.abspath(str(tmp_data_dir)) in result

    def test_path_inside_data_path_kept(self, tmp_data_dir):
        inside = os.path.join(str(tmp_data_dir), "subdir", "file.pptx")
        result = _resolve_data_path(inside)
        assert result == os.path.abspath(inside)


class TestSecureMakedirs:
    """Tests for _secure_makedirs()."""

    def test_creates_directory(self, tmp_path):
        new_dir = str(tmp_path / "secure_test")
        _secure_makedirs(new_dir)
        assert os.path.isdir(new_dir)

    def test_permissions_owner_only(self, tmp_path):
        new_dir = str(tmp_path / "perm_test")
        _secure_makedirs(new_dir)
        mode = os.stat(new_dir).st_mode & 0o777
        assert mode == 0o700

    def test_existing_directory_ok(self, tmp_path):
        existing = str(tmp_path / "existing")
        os.makedirs(existing)
        _secure_makedirs(existing)  # should not raise


class TestEnsureDirectory:
    """Tests for _ensure_directory()."""

    def test_creates_parent_dirs(self, tmp_path):
        file_path = str(tmp_path / "a" / "b" / "c" / "file.txt")
        result = _ensure_directory(file_path)
        assert os.path.isdir(os.path.dirname(result))

    def test_returns_absolute_path(self, tmp_path):
        result = _ensure_directory(str(tmp_path / "file.txt"))
        assert os.path.isabs(result)


class TestReadFileAsBase64:
    """Tests for _read_file_as_base64()."""

    def test_pptx_mime_type(self, pptx_file):
        result = _read_file_as_base64(pptx_file)
        assert result["mime_type"] == "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        assert result["file_name"] == os.path.basename(pptx_file)
        assert result["file_size_bytes"] > 0
        assert len(result["file_data_base64"]) > 0

    def test_xlsx_mime_type(self, xlsx_file):
        result = _read_file_as_base64(xlsx_file)
        assert result["mime_type"] == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    def test_docx_mime_type(self, docx_file):
        result = _read_file_as_base64(docx_file)
        assert result["mime_type"] == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

    def test_unknown_extension(self, tmp_data_dir):
        path = str(tmp_data_dir / "file.bin")
        with open(path, "wb") as f:
            f.write(b"binary data")
        result = _read_file_as_base64(path)
        assert result["mime_type"] == "application/octet-stream"
