"""MCP server exposing Microsoft Project desktop automation as tools."""

from __future__ import annotations

import argparse
import logging
import os
from typing import Any, Callable, Literal

from mcp.server.fastmcp import FastMCP
from mcp.server.fastmcp.exceptions import ToolError

from . import operations as ops
from .com_bridge import ComBridge, ProjectNotAvailableError

logger = logging.getLogger("msproject_mcp")

INSTRUCTIONS = """\
Controls Microsoft Project desktop on this Windows machine via COM automation.

- Tasks are addressed by their ID (the row number shown in Project). IDs shift
  when tasks are inserted or deleted, so re-read with list_tasks after
  structural changes. unique_id never changes.
- Durations are strings Project understands: "5d", "3h", "2w", "0d" (milestone).
- Dates are ISO 8601: "2026-10-05" or "2026-10-05T08:00".
- Resource units: 1.0 = 100%, 0.5 = 50%.
- Operations act on the active project unless project_name is given.
- Changes are not saved to disk until save_project is called.
"""


def _env_flag(name: str, default: bool) -> bool:
    value = os.environ.get(name)
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "on"}


def create_server(bridge: ComBridge | None = None) -> FastMCP:
    bridge = bridge or ComBridge(visible=_env_flag("MSPROJECT_VISIBLE", True))
    mcp = FastMCP("msproject", instructions=INSTRUCTIONS)

    async def call(fn: Callable[..., Any], **kwargs: Any) -> Any:
        try:
            return await bridge.run(lambda app: fn(app, **kwargs))
        except (ops.ProjectError, ProjectNotAvailableError) as exc:
            raise ToolError(str(exc)) from exc
        except Exception as exc:  # COM errors (pywintypes.com_error) and the like
            logger.exception("Project automation call %s failed", fn.__name__)
            raise ToolError(f"Microsoft Project error in {fn.__name__}: {_describe(exc)}") from exc

    # ---------------------------------------------------------------- app/files

    @mcp.tool()
    async def get_status() -> dict[str, Any]:
        """Connect to (or launch) Microsoft Project and list open projects."""
        return await call(ops.get_status)

    @mcp.tool()
    async def open_project(path: str, read_only: bool = False) -> dict[str, Any]:
        """Open a project file (.mpp, .mpt, .xml) and make it active.

        Args:
            path: Absolute path to the file on this machine.
            read_only: Open without write access.
        """
        return await call(ops.open_project, path=path, read_only=read_only)

    @mcp.tool()
    async def create_project(start_date: str | None = None, title: str | None = None) -> dict[str, Any]:
        """Create a new blank project and make it active.

        Args:
            start_date: Project start date (ISO), defaults to today.
            title: Name for the project summary task.
        """
        return await call(ops.create_project, start_date=start_date, title=title)

    @mcp.tool()
    async def activate_project(project_name: str) -> dict[str, Any]:
        """Switch the active project to another open project (by file name)."""
        return await call(ops.activate_project, project_name=project_name)

    @mcp.tool()
    async def save_project(
        path: str | None = None,
        file_format: Literal["mpp", "mpt", "xml"] = "mpp",
        project_name: str | None = None,
    ) -> dict[str, Any]:
        """Save the project. Without a path, saves in place; with a path, Save As.

        Args:
            path: Target file path for Save As (required for never-saved projects).
            file_format: mpp (default), mpt (template) or xml (MSPDI).
            project_name: Open project to save; defaults to the active one.
        """
        return await call(
            ops.save_project, path=path, file_format=file_format, project_name=project_name
        )

    @mcp.tool()
    async def export_pdf(path: str, project_name: str | None = None) -> dict[str, Any]:
        """Export the current view of a project to PDF."""
        return await call(ops.export_pdf, path=path, project_name=project_name)

    @mcp.tool()
    async def close_project(save: bool = False, project_name: str | None = None) -> dict[str, Any]:
        """Close a project. Unsaved changes are discarded unless save is true."""
        return await call(ops.close_project, save=save, project_name=project_name)

    @mcp.tool()
    async def get_project_info(project_name: str | None = None) -> dict[str, Any]:
        """Summary of a project: dates, duration, % complete, cost, counts."""
        return await call(ops.get_project_info, project_name=project_name)

    @mcp.tool()
    async def calculate_project(project_name: str | None = None) -> dict[str, Any]:
        """Recalculate the schedule of a project (the active one by default)."""
        return await call(ops.calculate_project, project_name=project_name)

    @mcp.tool()
    async def save_baseline(project_name: str | None = None) -> dict[str, Any]:
        """Save the current schedule of all tasks as the baseline."""
        return await call(ops.save_baseline, project_name=project_name)

    # -------------------------------------------------------------------- tasks

    @mcp.tool()
    async def list_tasks(
        name_contains: str | None = None,
        include_summary: bool = True,
        critical_only: bool = False,
        max_outline_level: int | None = None,
        offset: int = 0,
        limit: int = 200,
        project_name: str | None = None,
    ) -> dict[str, Any]:
        """List tasks with dates, duration, progress, links and resources.

        Args:
            name_contains: Case-insensitive substring filter on task name.
            include_summary: Include summary (parent) tasks.
            critical_only: Only tasks on the critical path.
            max_outline_level: Only tasks at this outline depth or shallower.
            offset: Skip this many matching tasks (pagination).
            limit: Maximum number of tasks to return.
        """
        return await call(
            ops.list_tasks,
            name_contains=name_contains,
            include_summary=include_summary,
            critical_only=critical_only,
            max_outline_level=max_outline_level,
            offset=offset,
            limit=limit,
            project_name=project_name,
        )

    @mcp.tool()
    async def get_task(task_id: int, project_name: str | None = None) -> dict[str, Any]:
        """Full details of one task, including notes, slack, baseline and assignments."""
        return await call(ops.get_task, task_id=task_id, project_name=project_name)

    @mcp.tool()
    async def add_task(
        name: str,
        duration: str | None = None,
        start: str | None = None,
        parent_id: int | None = None,
        before_id: int | None = None,
        predecessors: list[int] | None = None,
        notes: str | None = None,
        manually_scheduled: bool = False,
        project_name: str | None = None,
    ) -> dict[str, Any]:
        """Add a task. Appended at the end unless parent_id or before_id is given.

        Args:
            name: Task name.
            duration: e.g. "5d", "4h", "2w"; "0d" makes a milestone.
            start: Start date (ISO). Sets a constraint on auto-scheduled tasks.
            parent_id: Make the new task the last subtask of this task.
            before_id: Insert the new task above the task with this ID.
            predecessors: Task IDs to link as finish-to-start predecessors.
            notes: Task notes.
            manually_scheduled: Use manual scheduling instead of auto.
        """
        return await call(
            ops.add_task,
            name=name,
            duration=duration,
            start=start,
            parent_id=parent_id,
            before_id=before_id,
            predecessors=predecessors,
            notes=notes,
            manually_scheduled=manually_scheduled,
            project_name=project_name,
        )

    @mcp.tool()
    async def update_task(
        task_id: int,
        name: str | None = None,
        duration: str | None = None,
        start: str | None = None,
        finish: str | None = None,
        percent_complete: int | None = None,
        notes: str | None = None,
        priority: int | None = None,
        manually_scheduled: bool | None = None,
        active: bool | None = None,
        milestone: bool | None = None,
        deadline: str | None = None,
        constraint_type: Literal[
            "as_soon_as_possible",
            "as_late_as_possible",
            "must_start_on",
            "must_finish_on",
            "start_no_earlier_than",
            "start_no_later_than",
            "finish_no_earlier_than",
            "finish_no_later_than",
        ]
        | None = None,
        constraint_date: str | None = None,
        project_name: str | None = None,
    ) -> dict[str, Any]:
        """Update fields of a task. Only the fields you pass are changed.

        Args:
            task_id: ID of the task to update.
            percent_complete: 0-100.
            priority: 0-1000 (500 is the default).
            deadline: ISO date, or "" to clear the deadline.
            constraint_date: Required for must_*/start_*/finish_* constraints.
        """
        return await call(
            ops.update_task,
            task_id=task_id,
            name=name,
            duration=duration,
            start=start,
            finish=finish,
            percent_complete=percent_complete,
            notes=notes,
            priority=priority,
            manually_scheduled=manually_scheduled,
            active=active,
            milestone=milestone,
            deadline=deadline,
            constraint_type=constraint_type,
            constraint_date=constraint_date,
            project_name=project_name,
        )

    @mcp.tool()
    async def delete_task(task_id: int, project_name: str | None = None) -> dict[str, Any]:
        """Delete a task (and its subtasks, if it is a summary task)."""
        return await call(ops.delete_task, task_id=task_id, project_name=project_name)

    @mcp.tool()
    async def indent_task(task_id: int, project_name: str | None = None) -> dict[str, Any]:
        """Indent a task, making it a subtask of the task above it."""
        return await call(ops.indent_task, task_id=task_id, project_name=project_name)

    @mcp.tool()
    async def outdent_task(task_id: int, project_name: str | None = None) -> dict[str, Any]:
        """Outdent a task one outline level."""
        return await call(ops.outdent_task, task_id=task_id, project_name=project_name)

    @mcp.tool()
    async def link_tasks(
        predecessor_id: int,
        successor_id: int,
        link_type: Literal["FS", "SS", "FF", "SF"] = "FS",
        lag: str | None = None,
        project_name: str | None = None,
    ) -> dict[str, Any]:
        """Create a dependency between two tasks.

        Args:
            predecessor_id: Task that drives the successor.
            successor_id: Dependent task.
            link_type: FS (finish-to-start), SS, FF or SF.
            lag: Lag such as "2d"; negative for lead, e.g. "-1d"; or "50%".
        """
        return await call(
            ops.link_tasks,
            predecessor_id=predecessor_id,
            successor_id=successor_id,
            link_type=link_type,
            lag=lag,
            project_name=project_name,
        )

    @mcp.tool()
    async def unlink_tasks(
        predecessor_id: int, successor_id: int, project_name: str | None = None
    ) -> dict[str, Any]:
        """Remove the dependency between two tasks."""
        return await call(
            ops.unlink_tasks,
            predecessor_id=predecessor_id,
            successor_id=successor_id,
            project_name=project_name,
        )

    @mcp.tool()
    async def get_task_dependencies(task_id: int, project_name: str | None = None) -> dict[str, Any]:
        """List the predecessor and successor links of a task."""
        return await call(ops.get_task_dependencies, task_id=task_id, project_name=project_name)

    @mcp.tool()
    async def get_critical_path(project_name: str | None = None) -> dict[str, Any]:
        """List the non-summary tasks on the critical path."""
        return await call(ops.get_critical_path, project_name=project_name)

    # ---------------------------------------------------------------- resources

    @mcp.tool()
    async def list_resources(project_name: str | None = None) -> dict[str, Any]:
        """List resources with type, rate, max units, work, cost and overallocation."""
        return await call(ops.list_resources, project_name=project_name)

    @mcp.tool()
    async def add_resource(
        name: str,
        resource_type: Literal["work", "material", "cost"] = "work",
        max_units: float | None = None,
        standard_rate: str | None = None,
        initials: str | None = None,
        group: str | None = None,
        email: str | None = None,
        project_name: str | None = None,
    ) -> dict[str, Any]:
        """Add a resource to the project.

        Args:
            name: Resource name.
            resource_type: work (people/equipment), material or cost.
            max_units: Availability; 1.0 = 100%, 3.0 = three full-time people.
            standard_rate: Rate string such as "50/h" or "400/d".
        """
        return await call(
            ops.add_resource,
            name=name,
            resource_type=resource_type,
            max_units=max_units,
            standard_rate=standard_rate,
            initials=initials,
            group=group,
            email=email,
            project_name=project_name,
        )

    @mcp.tool()
    async def update_resource(
        resource: str,
        name: str | None = None,
        max_units: float | None = None,
        standard_rate: str | None = None,
        initials: str | None = None,
        group: str | None = None,
        email: str | None = None,
        project_name: str | None = None,
    ) -> dict[str, Any]:
        """Update a resource, identified by ID or exact name."""
        return await call(
            ops.update_resource,
            resource=resource,
            name=name,
            max_units=max_units,
            standard_rate=standard_rate,
            initials=initials,
            group=group,
            email=email,
            project_name=project_name,
        )

    @mcp.tool()
    async def delete_resource(resource: str, project_name: str | None = None) -> dict[str, Any]:
        """Delete a resource (by ID or name) and all of its assignments."""
        return await call(ops.delete_resource, resource=resource, project_name=project_name)

    @mcp.tool()
    async def assign_resource(
        task_id: int, resource: str, units: float = 1.0, project_name: str | None = None
    ) -> dict[str, Any]:
        """Assign a resource (by ID or name) to a task.

        Args:
            units: Assignment units; 1.0 = 100%, 0.5 = half time.
        """
        return await call(
            ops.assign_resource,
            task_id=task_id,
            resource=resource,
            units=units,
            project_name=project_name,
        )

    @mcp.tool()
    async def list_assignments(
        task_id: int | None = None, resource: str | None = None, project_name: str | None = None
    ) -> dict[str, Any]:
        """List assignments for a task, a resource, or the whole project."""
        return await call(
            ops.list_assignments, task_id=task_id, resource=resource, project_name=project_name
        )

    @mcp.tool()
    async def remove_assignment(
        task_id: int, resource: str, project_name: str | None = None
    ) -> dict[str, Any]:
        """Remove a resource assignment from a task."""
        return await call(
            ops.remove_assignment, task_id=task_id, resource=resource, project_name=project_name
        )

    return mcp


def _describe(exc: Exception) -> str:
    """Extract the readable message from a pywintypes.com_error."""
    args = getattr(exc, "args", ())
    # com_error args: (hresult, text, excepinfo, argerror)
    if len(args) >= 3 and isinstance(args[2], tuple) and len(args[2]) >= 3 and args[2][2]:
        return str(args[2][2])
    if len(args) >= 2 and isinstance(args[1], str):
        return args[1]
    return str(exc)


def main() -> None:
    parser = argparse.ArgumentParser(description="Microsoft Project MCP server")
    parser.add_argument(
        "--transport",
        choices=["stdio", "streamable-http", "sse"],
        default="stdio",
        help="MCP transport (default: stdio)",
    )
    parser.add_argument("--hidden", action="store_true", help="Run Project without showing its window")
    args = parser.parse_args()

    logging.basicConfig(level=os.environ.get("MSPROJECT_LOG_LEVEL", "WARNING"))
    visible = not args.hidden and _env_flag("MSPROJECT_VISIBLE", True)
    create_server(ComBridge(visible=visible)).run(transport=args.transport)


if __name__ == "__main__":
    main()
