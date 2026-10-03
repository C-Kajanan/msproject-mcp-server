# msproject-mcp-server

MCP server that lets Claude drive Microsoft Project desktop (Windows) via COM automation.

<!-- mcp-name: io.github.C-Kajanan/msproject-mcp-server -->

It attaches to a running copy of Project, or starts one, through the
`MSProject.Application` COM server. Claude can then open, build and edit
schedules: tasks, outline structure, dependencies, resources, assignments,
baselines, saving and PDF export.

## Requirements

- Windows with Microsoft Project desktop installed (Standard or Professional, 2016 or later)
- Python 3.10+

## Install

From PyPI (once a release is published):

```powershell
pip install msproject-mcp-server
```

Or run it without installing, using [uv](https://docs.astral.sh/uv/):

```powershell
uvx msproject-mcp-server
```

From source:

```powershell
git clone https://github.com/C-Kajanan/msproject-mcp-server.git
cd msproject-mcp-server
python -m venv .venv
.venv\Scripts\activate
pip install -e .
```

`pywin32` is installed automatically on Windows.

## Configure Claude

### Claude Desktop

Add this to `%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "msproject": {
      "command": "C:\\path\\to\\msproject-mcp-server\\.venv\\Scripts\\msproject-mcp.exe"
    }
  }
}
```

With uv, use `"command": "uvx", "args": ["msproject-mcp-server"]` instead.

### Claude Code

```powershell
claude mcp add msproject -- uvx msproject-mcp-server
# or, from a source checkout:
claude mcp add msproject -- C:\path\to\msproject-mcp-server\.venv\Scripts\msproject-mcp.exe
```

### Options

| Flag / env var | Effect |
| --- | --- |
| `--hidden` or `MSPROJECT_VISIBLE=0` | Keep the Project window hidden |
| `--transport streamable-http` | Serve over HTTP instead of stdio |
| `MSPROJECT_LOG_LEVEL=INFO` | More logging (on stderr) |

## Tools

| Area | Tools |
| --- | --- |
| Application and files | `get_status`, `open_project`, `create_project`, `activate_project`, `save_project` (mpp/mpt/xml), `export_pdf`, `close_project` |
| Project | `get_project_info`, `calculate_project`, `save_baseline`, `get_critical_path` |
| Tasks | `list_tasks`, `get_task`, `add_task`, `update_task`, `delete_task`, `indent_task`, `outdent_task` |
| Dependencies | `link_tasks` (FS/SS/FF/SF with lag), `unlink_tasks`, `get_task_dependencies` |
| Resources | `list_resources`, `add_resource`, `update_resource`, `delete_resource` |
| Assignments | `assign_resource`, `list_assignments`, `remove_assignment` |

Conventions:

- Tasks are addressed by **ID**, the row number shown in Project. IDs shift when rows are inserted or deleted; `unique_id` stays fixed.
- Durations and lags are strings Project understands: `"5d"`, `"4h"`, `"2w"`, `"0d"` (milestone), `"-1d"` (lead).
- Dates are ISO 8601: `2026-10-05` or `2026-10-05T08:00`.
- Units: `1.0` = 100%.
- Tools act on the active project unless you pass `project_name`.
- Nothing is written to disk until `save_project` is called.

Example prompt:

> Create a project starting 5 October 2026 for a website launch with phases Design, Build and Launch, add tasks with realistic durations, link them, assign Alice (50/h) and Bob (60/h), then save it to C:\Projects\website.mpp.

## Architecture

```
src/msproject_mcp/
  server.py       FastMCP tool definitions and CLI entry point
  operations.py   Project object-model logic; returns plain dicts
  com_bridge.py   Runs every COM call on one dedicated CoInitialize'd thread
  constants.py    Project enumeration values (link types, constraints...)
```

COM objects belong to the thread that created them, so all automation runs on
a single worker thread. The async MCP tools hand work to that thread, which
keeps the event loop free. If Project is closed while the server is running,
the next call reconnects.

## Development

The operations are tested against an in-memory fake of the Project object
model, so the tests run on any OS:

```bash
pip install -e ".[dev]"
pytest
```

## Releasing

Releases go to [PyPI](https://pypi.org/project/msproject-mcp-server/) and the
[MCP Registry](https://registry.modelcontextprotocol.io) as
`io.github.C-Kajanan/msproject-mcp-server`, via `.github/workflows/publish.yml`.

One-time setup on PyPI: add a trusted publisher at
<https://pypi.org/manage/account/publishing/> with project
`msproject-mcp-server`, owner `C-Kajanan`, repository `msproject-mcp-server`,
workflow `publish.yml` and environment `pypi`.

To release:

1. Bump `version` in `pyproject.toml` and both `version` fields in `server.json`.
2. Merge to `main`, then tag and push:

   ```bash
   git tag v0.1.0
   git push origin v0.1.0
   ```

The workflow checks the versions match the tag, runs the tests, publishes to
PyPI, then publishes `server.json` to the MCP Registry using GitHub OIDC (no
token needed).

## Limitations

- Windows only. COM automation does not work against Project for the web or Project Online.
- Project shows a modal dialog for some actions (for example, a planning wizard warning). The server sets `DisplayAlerts = False` to suppress most of them. If a call hangs, look at the Project window.
