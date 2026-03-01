"""Tests for CLI entry point."""

import os
import signal
import sys
from unittest.mock import patch, MagicMock

import pytest

from onit_office.cli import (
    _read_pid,
    _write_pid,
    cmd_start,
    cmd_stop,
    cmd_status,
    main,
)


@pytest.fixture
def cli_dir(tmp_path, monkeypatch):
    """Redirect CLI paths to tmp directory."""
    import onit_office.cli as cli_mod
    monkeypatch.setattr(cli_mod, "ONIT_DIR", str(tmp_path))
    monkeypatch.setattr(cli_mod, "PID_FILE", str(tmp_path / "server.pid"))
    monkeypatch.setattr(cli_mod, "LOG_FILE", str(tmp_path / "server.log"))
    return tmp_path


class TestReadPid:
    """Tests for _read_pid()."""

    def test_no_pid_file(self, cli_dir):
        assert _read_pid() is None

    def test_stale_pid(self, cli_dir):
        # Write a PID that doesn't exist
        _write_pid(999999)
        assert _read_pid() is None

    def test_valid_pid(self, cli_dir):
        # Current process PID should be alive
        _write_pid(os.getpid())
        assert _read_pid() == os.getpid()

    def test_invalid_pid_content(self, cli_dir):
        import onit_office.cli as cli_mod
        with open(cli_mod.PID_FILE, "w") as f:
            f.write("not_a_number")
        assert _read_pid() is None


class TestWritePid:
    """Tests for _write_pid()."""

    def test_writes_pid(self, cli_dir):
        _write_pid(12345)
        import onit_office.cli as cli_mod
        with open(cli_mod.PID_FILE) as f:
            assert f.read().strip() == "12345"

    def test_creates_directory(self, tmp_path, monkeypatch):
        import onit_office.cli as cli_mod
        new_dir = str(tmp_path / "new_dir")
        monkeypatch.setattr(cli_mod, "ONIT_DIR", new_dir)
        monkeypatch.setattr(cli_mod, "PID_FILE", os.path.join(new_dir, "server.pid"))
        _write_pid(12345)
        assert os.path.isdir(new_dir)


class TestCmdStatus:
    """Tests for cmd_status()."""

    def test_not_running(self, cli_dir, capsys):
        args = MagicMock()
        cmd_status(args)
        captured = capsys.readouterr()
        assert "not running" in captured.out

    def test_running(self, cli_dir, capsys):
        _write_pid(os.getpid())
        args = MagicMock()
        cmd_status(args)
        captured = capsys.readouterr()
        assert "is running" in captured.out
        assert str(os.getpid()) in captured.out


class TestCmdStart:
    """Tests for cmd_start()."""

    def test_already_running(self, cli_dir, capsys):
        _write_pid(os.getpid())
        args = MagicMock()
        args.foreground = False
        cmd_start(args)
        captured = capsys.readouterr()
        assert "already running" in captured.out


class TestCmdStop:
    """Tests for cmd_stop()."""

    def test_not_running(self, cli_dir, capsys):
        args = MagicMock()
        cmd_stop(args)
        captured = capsys.readouterr()
        assert "not running" in captured.out


class TestMain:
    """Tests for main() argument parsing."""

    def test_version(self, capsys):
        with pytest.raises(SystemExit) as exc_info:
            with patch("sys.argv", ["onit-office", "--version"]):
                main()
        assert exc_info.value.code == 0

    def test_status_command(self, cli_dir, capsys):
        with patch("sys.argv", ["onit-office", "status"]):
            main()
        captured = capsys.readouterr()
        assert "not running" in captured.out

    def test_stop_command(self, cli_dir, capsys):
        with patch("sys.argv", ["onit-office", "stop"]):
            main()
        captured = capsys.readouterr()
        assert "not running" in captured.out
