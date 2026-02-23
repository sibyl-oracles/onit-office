"""SSE server runner for onit-office MCP server."""

import os
import sys

from .mcp_server import mcp, _secure_makedirs


def start_server(
    host: str = "0.0.0.0",
    port: int = 18203,
    data_path: str = "",
) -> None:
    """Start the MCP server with SSE transport."""
    from . import mcp_server

    if data_path:
        mcp_server.DATA_PATH = data_path

    _secure_makedirs(os.path.abspath(os.path.expanduser(mcp_server.DATA_PATH)))

    print(f"Starting onit-office MCP server (SSE) on http://{host}:{port}/sse")
    print(f"Data path: {mcp_server.DATA_PATH}")
    sys.stdout.flush()

    mcp.run(transport="sse", host=host, port=port)


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("--host", default="0.0.0.0")
    parser.add_argument("--port", type=int, default=18203)
    parser.add_argument("--data-path", default="")
    args = parser.parse_args()

    start_server(host=args.host, port=args.port, data_path=args.data_path)
