"""Entry point for the Claude Desktop extension (MCPB bundle).

The host runs this file as a script with uv, so the package directory next
to it is put on sys.path explicitly instead of relying on the project being
installed into the environment.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from outlook_desktop_mcp.entrypoint import main  # noqa: E402

if __name__ == "__main__":
    main()
