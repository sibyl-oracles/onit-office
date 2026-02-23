"""CLI entry point for onit-office MCP server."""

import argparse
import os
import signal
import subprocess
import sys
import time

ONIT_DIR = os.path.join(os.path.expanduser("~"), ".onit-office")
PID_FILE = os.path.join(ONIT_DIR, "server.pid")
LOG_FILE = os.path.join(ONIT_DIR, "server.log")


def _read_pid() -> int | None:
    """Read PID from file, return None if not found or stale."""
    if not os.path.exists(PID_FILE):
        return None
    try:
        with open(PID_FILE) as f:
            pid = int(f.read().strip())
        # Check if process is alive
        os.kill(pid, 0)
        return pid
    except (ValueError, ProcessLookupError, PermissionError, OSError):
        # Stale PID file
        try:
            os.remove(PID_FILE)
        except OSError:
            pass
        return None


def _write_pid(pid: int) -> None:
    os.makedirs(ONIT_DIR, mode=0o700, exist_ok=True)
    with open(PID_FILE, "w") as f:
        f.write(str(pid))


def cmd_start(args: argparse.Namespace) -> None:
    """Start the MCP server."""
    pid = _read_pid()
    if pid is not None:
        print(f"onit-office is already running (PID {pid})")
        return

    if args.foreground:
        # Run in foreground (blocking)
        from .server import start_server
        start_server(host=args.host, port=args.port, data_path=args.data_path)
        return

    # Run in background
    os.makedirs(ONIT_DIR, mode=0o700, exist_ok=True)

    log = open(LOG_FILE, "a")
    proc = subprocess.Popen(
        [
            sys.executable, "-m", "onit_office.server",
            "--host", args.host,
            "--port", str(args.port),
            "--data-path", args.data_path,
        ],
        stdout=log,
        stderr=log,
        start_new_session=True,
    )

    _write_pid(proc.pid)

    # Brief wait to check it didn't crash immediately
    time.sleep(1)
    if proc.poll() is not None:
        print("onit-office failed to start. Check logs:")
        print(f"  {LOG_FILE}")
        try:
            os.remove(PID_FILE)
        except OSError:
            pass
        return

    print(f"onit-office started (PID {proc.pid})")
    print(f"  SSE endpoint: http://{args.host}:{args.port}/sse")
    print(f"  Log file: {LOG_FILE}")


def cmd_stop(args: argparse.Namespace) -> None:
    """Stop the MCP server."""
    pid = _read_pid()
    if pid is None:
        print("onit-office is not running")
        return

    try:
        os.kill(pid, signal.SIGTERM)
        # Wait for process to exit
        for _ in range(30):
            time.sleep(0.1)
            try:
                os.kill(pid, 0)
            except ProcessLookupError:
                break
        else:
            # Force kill if still running
            os.kill(pid, signal.SIGKILL)
    except ProcessLookupError:
        pass

    try:
        os.remove(PID_FILE)
    except OSError:
        pass

    print(f"onit-office stopped (PID {pid})")


def cmd_status(args: argparse.Namespace) -> None:
    """Check server status."""
    pid = _read_pid()
    if pid is not None:
        print(f"onit-office is running (PID {pid})")
    else:
        print("onit-office is not running")


def main() -> None:
    parser = argparse.ArgumentParser(
        prog="onit-office",
        description="Standalone MCP server for Microsoft Office document creation",
    )
    parser.add_argument(
        "--version", action="version",
        version=f"%(prog)s {__import__('onit_office').__version__}",
    )

    subparsers = parser.add_subparsers(dest="command")

    # start command
    start_parser = subparsers.add_parser("start", help="Start the server")
    start_parser.add_argument("--host", default="0.0.0.0", help="Host to bind (default: 0.0.0.0)")
    start_parser.add_argument("--port", type=int, default=18203, help="Port to bind (default: 18203)")
    start_parser.add_argument("--data-path", default="", help="Data directory for created files")
    start_parser.add_argument("--foreground", action="store_true", help="Run in foreground (don't daemonize)")

    # stop command
    subparsers.add_parser("stop", help="Stop the server")

    # status command
    subparsers.add_parser("status", help="Check server status")

    args = parser.parse_args()

    # Default to "start" if no subcommand given
    if args.command is None:
        # Re-parse with "start" defaults
        args = start_parser.parse_args(sys.argv[1:])
        args.command = "start"

    if args.command == "start":
        cmd_start(args)
    elif args.command == "stop":
        cmd_stop(args)
    elif args.command == "status":
        cmd_status(args)


if __name__ == "__main__":
    main()
