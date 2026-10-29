"""Build the Claude Desktop extension (dist/outlook-desktop-mcp-<version>.mcpb).

Regenerates manifest.json from pyproject.toml (version, dependencies) and
from the @mcp.tool functions in server.py (tool list), then validates and
packs the bundle with the mcpb CLI (needs Node.js for npx).

Run: python scripts/build_mcpb.py
"""
import ast
import json
import shutil
import subprocess
import sys
import tomllib
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SERVER = ROOT / "src" / "outlook_desktop_mcp" / "server.py"
MCPB = ["npx", "-y", "@anthropic-ai/mcpb"]


def _tools() -> list[dict]:
    """Name and first docstring line of every @mcp.tool() function."""
    tools = []
    for node in ast.parse(SERVER.read_text(encoding="utf-8")).body:
        if not isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            continue
        if not any(ast.unparse(d) == "mcp.tool()" for d in node.decorator_list):
            continue
        doc = ast.get_docstring(node) or ""
        tools.append({"name": node.name, "description": doc.split("\n\n")[0].replace("\n", " ")})
    return tools


def build_manifest() -> dict:
    project = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))["project"]
    author = project["authors"][0]
    return {
        "manifest_version": "0.4",
        "name": project["name"],
        "display_name": "Outlook Desktop",
        "version": project["version"],
        "description": "Email, calendar, tasks and To-Do in your local Outlook — no Graph API, no app registration.",
        "long_description": (
            "Drives the Outlook Desktop app already signed in on this computer: "
            "Classic Outlook on Windows through COM, Outlook for Mac through "
            "AppleScript. Outlook must be running. The new Outlook (olk.exe) "
            "is not supported."
        ),
        "author": {"name": author["name"]},
        "repository": {"type": "git", "url": "https://github.com/iu3qez/outlook-desktop-mcp"},
        "license": "AGPL-3.0",
        "keywords": project["keywords"],
        "server": {
            "type": "uv",
            "entry_point": "src/mcpb_server.py",
            "mcp_config": {
                "command": "uv",
                "args": ["run", "--directory", "${__dirname}", "src/mcpb_server.py"],
            },
        },
        "tools": _tools(),
        "compatibility": {
            "platforms": ["win32", "darwin"],
            "runtimes": {"python": project["requires-python"]},
        },
    }


def main():
    manifest = build_manifest()
    (ROOT / "manifest.json").write_text(json.dumps(manifest, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    print(f"manifest.json: {manifest['version']}, {len(manifest['tools'])} tools")

    if shutil.which("npx") is None:
        sys.exit("npx not found: install Node.js to pack the bundle")
    out = ROOT / "dist" / f"{manifest['name']}-{manifest['version']}.mcpb"
    out.parent.mkdir(exist_ok=True)
    subprocess.run([*MCPB, "validate", "manifest.json"], cwd=ROOT, check=True, shell=sys.platform == "win32")
    subprocess.run([*MCPB, "pack", ".", str(out)], cwd=ROOT, check=True, shell=sys.platform == "win32")


if __name__ == "__main__":
    main()
