"""Shared fixtures for onit-office tests."""

import os
import json
import shutil
import tempfile

import pytest


@pytest.fixture
def tmp_data_dir(tmp_path):
    """Provide a temporary data directory and patch DATA_PATH."""
    from onit_office import mcp_server

    original = mcp_server.DATA_PATH
    mcp_server.DATA_PATH = str(tmp_path)
    os.makedirs(str(tmp_path), exist_ok=True)
    yield tmp_path
    mcp_server.DATA_PATH = original


@pytest.fixture
def pptx_file(tmp_data_dir):
    """Create a presentation and return its path."""
    from onit_office.mcp_server import create_presentation

    result = json.loads(create_presentation(
        title="Test Presentation",
        subtitle="Test Subtitle",
        path=str(tmp_data_dir / "test.pptx"),
    ))
    assert "error" not in result
    return result["powerpoint_file"]


@pytest.fixture
def xlsx_file(tmp_data_dir):
    """Create an Excel file and return its path."""
    from onit_office.mcp_server import create_excel

    result = json.loads(create_excel(
        path=str(tmp_data_dir / "test.xlsx"),
        headers=["Name", "Age", "Email"],
        rows=[["Alice", 30, "alice@example.com"], ["Bob", 25, "bob@example.com"]],
    ))
    assert "error" not in result
    return result["excel_file"]


@pytest.fixture
def docx_file(tmp_data_dir):
    """Create a Word document and return its path."""
    from onit_office.mcp_server import create_document

    result = json.loads(create_document(
        path=str(tmp_data_dir / "test.docx"),
        title="Test Document",
    ))
    assert "error" not in result
    return result["document_file"]


@pytest.fixture
def sample_image(tmp_data_dir):
    """Create a small test PNG image and return its path."""
    from PIL import Image

    img_path = str(tmp_data_dir / "test_image.png")
    img = Image.new("RGB", (100, 100), color="red")
    img.save(img_path)
    return img_path
