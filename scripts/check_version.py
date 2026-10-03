"""Fail if pyproject.toml, server.json, plugin.json and (optionally) a release tag disagree on the version."""

import json
import pathlib
import sys
import tomllib

root = pathlib.Path(__file__).resolve().parent.parent
version = tomllib.loads((root / "pyproject.toml").read_text())["project"]["version"]
server = json.loads((root / "server.json").read_text())
plugin = json.loads((root / ".claude-plugin" / "plugin.json").read_text())

errors = []
if plugin["version"] != version:
    errors.append(f"plugin.json version {plugin['version']} != pyproject {version}")
if server["version"] != version:
    errors.append(f"server.json version {server['version']} != pyproject {version}")
for pkg in server["packages"]:
    if pkg["version"] != version:
        errors.append(f"server.json package version {pkg['version']} != pyproject {version}")
if len(sys.argv) > 1:
    tag = sys.argv[1].removeprefix("refs/tags/").removeprefix("v")
    if tag != version:
        errors.append(f"tag v{tag} != pyproject {version}")

if errors:
    sys.exit("\n".join(errors))
print(f"version {version} is consistent")
