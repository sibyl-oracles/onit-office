"""Tests for get_file and download_media tools in mcp_server.py."""

import base64
import json
import os
from unittest.mock import patch, MagicMock

import pytest

from onit_office.mcp_server import get_file, download_media


class TestGetFile:
    """Tests for get_file()."""

    def test_get_pptx(self, pptx_file):
        result = json.loads(get_file(path=pptx_file))
        assert result["status"] == "success"
        assert result["file_name"] == os.path.basename(pptx_file)
        assert result["mime_type"] == "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        assert result["file_size_bytes"] > 0
        # Verify base64 is valid
        decoded = base64.b64decode(result["file_data_base64"])
        assert len(decoded) == result["file_size_bytes"]

    def test_get_xlsx(self, xlsx_file):
        result = json.loads(get_file(path=xlsx_file))
        assert result["status"] == "success"
        assert result["mime_type"] == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    def test_get_docx(self, docx_file):
        result = json.loads(get_file(path=docx_file))
        assert result["status"] == "success"
        assert result["mime_type"] == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

    def test_nonexistent_file(self, tmp_data_dir):
        result = json.loads(get_file(path=str(tmp_data_dir / "nope.pptx")))
        assert "error" in result
        assert result["status"] == "failed"

    def test_relative_path_resolved(self, tmp_data_dir):
        """A relative path should be resolved under DATA_PATH."""
        from onit_office.mcp_server import create_presentation
        create_presentation(title="Rel", path=str(tmp_data_dir / "rel.pptx"))
        result = json.loads(get_file(path="rel.pptx"))
        assert result["status"] == "success"


class TestDownloadMedia:
    """Tests for download_media()."""

    def test_invalid_url_scheme(self, tmp_data_dir):
        result = json.loads(download_media(
            url="ftp://example.com/file.png",
            output_path=str(tmp_data_dir / "file.png"),
        ))
        assert "error" in result
        assert "http" in result["error"]

    @patch("onit_office.mcp_server.requests.get")
    def test_successful_download_png(self, mock_get, tmp_data_dir):
        # Create a minimal PNG header
        png_header = b'\x89PNG\r\n\x1a\n' + b'\x00' * 100
        mock_response = MagicMock()
        mock_response.content = png_header
        mock_response.headers = {"Content-Type": "image/png"}
        mock_response.raise_for_status = MagicMock()
        mock_get.return_value = mock_response

        result = json.loads(download_media(
            url="https://example.com/image.png",
            output_path=str(tmp_data_dir / "downloaded.png"),
        ))
        assert result["media_type"] == "image"
        assert result["content_type"] == "image/png"
        assert result["size_bytes"] == len(png_header)
        assert os.path.exists(result["file_path"])

    @patch("onit_office.mcp_server.requests.get")
    def test_successful_download_jpeg(self, mock_get, tmp_data_dir):
        jpeg_header = b'\xff\xd8\xff' + b'\x00' * 50
        mock_response = MagicMock()
        mock_response.content = jpeg_header
        mock_response.headers = {"Content-Type": "image/jpeg"}
        mock_response.raise_for_status = MagicMock()
        mock_get.return_value = mock_response

        result = json.loads(download_media(
            url="https://example.com/photo.jpg",
            output_path=str(tmp_data_dir / "photo.jpg"),
        ))
        assert result["media_type"] == "image"
        assert result["content_type"] == "image/jpeg"

    @patch("onit_office.mcp_server.requests.get")
    def test_successful_download_pdf(self, mock_get, tmp_data_dir):
        pdf_content = b'%PDF-1.4 test content'
        mock_response = MagicMock()
        mock_response.content = pdf_content
        mock_response.headers = {"Content-Type": "application/pdf"}
        mock_response.raise_for_status = MagicMock()
        mock_get.return_value = mock_response

        result = json.loads(download_media(
            url="https://example.com/doc.pdf",
            output_path=str(tmp_data_dir / "doc.pdf"),
        ))
        assert result["media_type"] == "document"

    @patch("onit_office.mcp_server.requests.get")
    def test_successful_download_mp3(self, mock_get, tmp_data_dir):
        mp3_content = b'ID3' + b'\x00' * 50
        mock_response = MagicMock()
        mock_response.content = mp3_content
        mock_response.headers = {"Content-Type": "audio/mpeg"}
        mock_response.raise_for_status = MagicMock()
        mock_get.return_value = mock_response

        result = json.loads(download_media(
            url="https://example.com/song.mp3",
            output_path=str(tmp_data_dir / "song.mp3"),
        ))
        assert result["media_type"] == "audio"

    @patch("onit_office.mcp_server.requests.get")
    def test_download_timeout(self, mock_get, tmp_data_dir):
        import requests as req
        mock_get.side_effect = req.exceptions.Timeout("Connection timed out")

        result = json.loads(download_media(
            url="https://example.com/slow.png",
            output_path=str(tmp_data_dir / "slow.png"),
            timeout=5,
        ))
        assert "error" in result
        assert "timed out" in result["error"]

    @patch("onit_office.mcp_server.requests.get")
    def test_download_request_error(self, mock_get, tmp_data_dir):
        import requests as req
        mock_get.side_effect = req.exceptions.ConnectionError("Connection refused")

        result = json.loads(download_media(
            url="https://example.com/fail.png",
            output_path=str(tmp_data_dir / "fail.png"),
        ))
        assert "error" in result

    @patch("onit_office.mcp_server.requests.get")
    def test_fallback_to_content_type_header(self, mock_get, tmp_data_dir):
        # Content with no recognized signature
        mock_response = MagicMock()
        mock_response.content = b'unknown binary content'
        mock_response.headers = {"Content-Type": "video/mp4"}
        mock_response.raise_for_status = MagicMock()
        mock_get.return_value = mock_response

        result = json.loads(download_media(
            url="https://example.com/video.mp4",
            output_path=str(tmp_data_dir / "video.mp4"),
        ))
        assert result["media_type"] == "video"

    @patch("onit_office.mcp_server.requests.get")
    def test_file_permissions(self, mock_get, tmp_data_dir):
        mock_response = MagicMock()
        mock_response.content = b'\x89PNG\r\n\x1a\n' + b'\x00' * 10
        mock_response.headers = {"Content-Type": "image/png"}
        mock_response.raise_for_status = MagicMock()
        mock_get.return_value = mock_response

        output = str(tmp_data_dir / "perms.png")
        download_media(url="https://example.com/img.png", output_path=output)
        mode = os.stat(output).st_mode & 0o777
        assert mode == 0o600
